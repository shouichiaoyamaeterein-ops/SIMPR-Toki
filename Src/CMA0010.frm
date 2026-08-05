VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~1.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Begin VB.Form CMA0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "受注納品書入力 [丸高]"
   ClientHeight    =   10980
   ClientLeft      =   390
   ClientTop       =   975
   ClientWidth     =   15330
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
   ScaleWidth      =   15330
   WindowState     =   2  '最大化
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   630
      ScaleHeight     =   1890
      ScaleWidth      =   3675
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   3435
      Width           =   3735
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1725
         Left            =   90
         Top             =   90
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   3043
         ForeColor       =   16711680
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         cAlingnment     =   6
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   840
            Index           =   3
            Left            =   1170
            Top             =   30
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   1482
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
            cPositionX      =   100
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   345
               Left            =   1185
               Top             =   15
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "変換先ｺｰﾄﾞ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ Ｐゴシック"
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
               cAlingnment     =   1
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   300
               Left            =   855
               Top             =   405
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   529
               Caption         =   "→"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   15.75
                  Charset         =   128
                  Weight          =   700
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
            Begin Cis3D_v60.CIS3D Blb_HCord 
               Height          =   375
               Left            =   1260
               Top             =   360
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
               BackColor       =   16777152
               Caption         =   "XXXXXXX"
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
            End
            Begin CisText_V60.CisText iB1_NTorcd 
               Height          =   360
               Left            =   90
               TabIndex        =   11
               Top             =   360
               Width           =   735
               _ExtentX        =   1296
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   840
            Index           =   2
            Left            =   2430
            Top             =   870
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   1482
            ForeColor       =   16711680
            Caption         =   "受注者"
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
            Begin CisText_V60.CisText iB1_JTorcd 
               Height          =   375
               Left            =   165
               TabIndex        =   14
               Top             =   375
               Width           =   720
               _ExtentX        =   1270
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
               Text            =   "XXXXX"
               MaxLength       =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   840
            Index           =   0
            Left            =   30
            Top             =   30
            Width           =   1140
            _ExtentX        =   2011
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
            cPositionY      =   30
            Begin CisText_V60.CisText iB1_Noubn 
               Height          =   375
               Left            =   195
               TabIndex        =   10
               Top             =   375
               Width           =   750
               _ExtentX        =   1323
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
               Text            =   "XXXXX"
               MaxLength       =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   840
            Index           =   12
            Left            =   30
            Top             =   870
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   1482
            ForeColor       =   16711680
            Caption         =   "納入指示日-便"
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
            Begin CisYMD_With_Btn.CisYMDwB iB1_NSYmd 
               Height          =   360
               Left            =   75
               TabIndex        =   12
               Top             =   360
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
            Begin CisText_V60.CisText iB1_NBin 
               Height          =   375
               Left            =   1935
               TabIndex        =   13
               Top             =   360
               Width           =   345
               _ExtentX        =   609
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
               Text            =   "99"
               MaxLength       =   2
               cDataType       =   1
               cDataReplace    =   1
               cFaZero         =   0
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
         End
      End
   End
   Begin VB.PictureBox B1_Area2 
      BackColor       =   &H00C0C0C0&
      Height          =   8895
      Left            =   6390
      ScaleHeight     =   8835
      ScaleWidth      =   7410
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   1410
      Width           =   7470
      Begin VB.PictureBox B1_Area9 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8745
         Left            =   75
         ScaleHeight     =   8685
         ScaleWidth      =   7215
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   45
         Width           =   7275
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   360
            Left            =   0
            Top             =   0
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "品     番"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D UC_3D23 
            Height          =   360
            Left            =   3840
            Top             =   0
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   635
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
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D UC_3D21 
            Height          =   360
            Left            =   4965
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "箱数"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D UC_3D18 
            Height          =   360
            Left            =   2955
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   635
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
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D UC_3D26 
            Height          =   360
            Left            =   5850
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "納入数"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   0
            Left            =   4965
            Top             =   360
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   0
               Left            =   210
               TabIndex        =   18
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   0
            Left            =   3840
            Top             =   360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   0
               Left            =   225
               TabIndex        =   17
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   0
            Left            =   2955
            Top             =   360
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   0
               Left            =   120
               TabIndex        =   16
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   0
            Left            =   0
            Top             =   360
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   0
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   0
               Left            =   525
               TabIndex        =   15
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   0
            Left            =   5850
            Top             =   360
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   0
               Left            =   210
               TabIndex        =   19
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   1
            Left            =   4965
            Top             =   915
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   1
               Left            =   210
               TabIndex        =   23
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   1
            Left            =   3840
            Top             =   915
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   1
               Left            =   225
               TabIndex        =   22
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   1
            Left            =   2955
            Top             =   915
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   1
               Left            =   120
               TabIndex        =   21
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   1
            Left            =   0
            Top             =   915
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   1
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   1
               Left            =   525
               TabIndex        =   20
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   1
            Left            =   5850
            Top             =   915
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   1
               Left            =   210
               TabIndex        =   24
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   2
            Left            =   4965
            Top             =   1470
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   2
               Left            =   210
               TabIndex        =   28
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   2
            Left            =   3840
            Top             =   1470
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   2
               Left            =   225
               TabIndex        =   27
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   2
            Left            =   2955
            Top             =   1470
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   2
               Left            =   120
               TabIndex        =   26
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   2
            Left            =   0
            Top             =   1470
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   2
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   2
               Left            =   525
               TabIndex        =   25
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   2
            Left            =   5850
            Top             =   1470
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   2
               Left            =   210
               TabIndex        =   29
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   3
            Left            =   4965
            Top             =   2025
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   3
               Left            =   210
               TabIndex        =   33
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   3
            Left            =   3840
            Top             =   2025
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   3
               Left            =   225
               TabIndex        =   32
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   3
            Left            =   2955
            Top             =   2025
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   3
               Left            =   120
               TabIndex        =   31
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   3
            Left            =   0
            Top             =   2025
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   3
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   3
               Left            =   525
               TabIndex        =   30
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   3
            Left            =   5850
            Top             =   2025
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   3
               Left            =   210
               TabIndex        =   34
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   4
            Left            =   4965
            Top             =   2580
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   4
               Left            =   210
               TabIndex        =   38
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   4
            Left            =   3840
            Top             =   2580
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   4
               Left            =   225
               TabIndex        =   37
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   4
            Left            =   2955
            Top             =   2580
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   4
               Left            =   120
               TabIndex        =   36
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   4
            Left            =   0
            Top             =   2580
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   4
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   4
               Left            =   525
               TabIndex        =   35
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   4
            Left            =   5850
            Top             =   2580
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   4
               Left            =   210
               TabIndex        =   39
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   5
            Left            =   4965
            Top             =   3135
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   5
               Left            =   210
               TabIndex        =   43
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   5
            Left            =   3840
            Top             =   3135
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   5
               Left            =   225
               TabIndex        =   42
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   5
            Left            =   2955
            Top             =   3135
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   5
               Left            =   120
               TabIndex        =   41
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   5
            Left            =   0
            Top             =   3135
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   5
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   5
               Left            =   525
               TabIndex        =   40
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   5
            Left            =   5850
            Top             =   3135
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   5
               Left            =   210
               TabIndex        =   44
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   6
            Left            =   4965
            Top             =   3690
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   6
               Left            =   210
               TabIndex        =   48
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   6
            Left            =   3840
            Top             =   3690
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   6
               Left            =   225
               TabIndex        =   47
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   6
            Left            =   2955
            Top             =   3690
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   6
               Left            =   120
               TabIndex        =   46
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   6
            Left            =   0
            Top             =   3690
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   6
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   6
               Left            =   525
               TabIndex        =   45
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   6
            Left            =   5850
            Top             =   3690
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   6
               Left            =   210
               TabIndex        =   49
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   7
            Left            =   4965
            Top             =   4245
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   7
               Left            =   210
               TabIndex        =   53
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   7
            Left            =   3840
            Top             =   4245
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   7
               Left            =   225
               TabIndex        =   52
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   7
            Left            =   2955
            Top             =   4245
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   7
               Left            =   120
               TabIndex        =   51
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   7
            Left            =   0
            Top             =   4245
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   7
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   7
               Left            =   525
               TabIndex        =   50
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   7
            Left            =   5850
            Top             =   4245
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   7
               Left            =   210
               TabIndex        =   54
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   8
            Left            =   4965
            Top             =   4800
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   8
               Left            =   210
               TabIndex        =   58
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   8
            Left            =   3840
            Top             =   4800
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   8
               Left            =   225
               TabIndex        =   57
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   8
            Left            =   2955
            Top             =   4800
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   8
               Left            =   120
               TabIndex        =   56
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   8
            Left            =   0
            Top             =   4800
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   8
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   8
               Left            =   525
               TabIndex        =   55
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   8
            Left            =   5850
            Top             =   4800
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   8
               Left            =   210
               TabIndex        =   59
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   9
            Left            =   4965
            Top             =   5355
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   9
               Left            =   210
               TabIndex        =   63
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   9
            Left            =   3840
            Top             =   5355
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   9
               Left            =   225
               TabIndex        =   62
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   9
            Left            =   2955
            Top             =   5355
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   9
               Left            =   120
               TabIndex        =   61
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   9
            Left            =   0
            Top             =   5355
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   9
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   9
               Left            =   525
               TabIndex        =   60
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   9
            Left            =   5850
            Top             =   5355
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   9
               Left            =   210
               TabIndex        =   64
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   10
            Left            =   4965
            Top             =   5910
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   10
               Left            =   210
               TabIndex        =   68
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   10
            Left            =   3840
            Top             =   5910
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   10
               Left            =   225
               TabIndex        =   67
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   10
            Left            =   2955
            Top             =   5910
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   10
               Left            =   120
               TabIndex        =   66
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   10
            Left            =   0
            Top             =   5910
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   10
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "11"
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
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   10
               Left            =   525
               TabIndex        =   65
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   10
            Left            =   5850
            Top             =   5910
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   10
               Left            =   210
               TabIndex        =   69
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   11
            Left            =   4965
            Top             =   6465
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   11
               Left            =   210
               TabIndex        =   73
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   11
            Left            =   3840
            Top             =   6465
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   11
               Left            =   225
               TabIndex        =   72
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   11
            Left            =   2955
            Top             =   6465
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   11
               Left            =   120
               TabIndex        =   71
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   11
            Left            =   0
            Top             =   6465
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   11
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "12"
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
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   11
               Left            =   525
               TabIndex        =   70
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   11
            Left            =   5850
            Top             =   6465
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   11
               Left            =   210
               TabIndex        =   74
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   12
            Left            =   4965
            Top             =   7020
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   12
               Left            =   210
               TabIndex        =   78
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   12
            Left            =   3840
            Top             =   7020
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   12
               Left            =   225
               TabIndex        =   77
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   12
            Left            =   2955
            Top             =   7020
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   12
               Left            =   120
               TabIndex        =   76
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   12
            Left            =   0
            Top             =   7020
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   12
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "13"
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
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   12
               Left            =   525
               TabIndex        =   75
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   12
            Left            =   5850
            Top             =   7020
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   12
               Left            =   210
               TabIndex        =   79
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   13
            Left            =   4965
            Top             =   7575
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   13
               Left            =   210
               TabIndex        =   83
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   13
            Left            =   3840
            Top             =   7575
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   13
               Left            =   225
               TabIndex        =   82
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   13
            Left            =   2955
            Top             =   7575
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   13
               Left            =   120
               TabIndex        =   81
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   13
            Left            =   0
            Top             =   7575
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   13
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "14"
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
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   13
               Left            =   525
               TabIndex        =   80
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   13
            Left            =   5850
            Top             =   7575
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   13
               Left            =   210
               TabIndex        =   84
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   14
            Left            =   4965
            Top             =   8130
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   14
               Left            =   210
               TabIndex        =   88
               Top             =   105
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   14
            Left            =   3840
            Top             =   8130
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   14
               Left            =   225
               TabIndex        =   87
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   635
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   " 9999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   14
            Left            =   2955
            Top             =   8130
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   14
               Left            =   120
               TabIndex        =   86
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   14
            Left            =   0
            Top             =   8130
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   14
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "15"
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
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   14
               Left            =   525
               TabIndex        =   85
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XX"
               MaxLength       =   12
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   14
            Left            =   5850
            Top             =   8130
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   14
               Left            =   210
               TabIndex        =   89
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               cFormat         =   "#######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
      End
   End
   Begin PdqcommLib.PDQComm PDQComm1 
      Height          =   480
      Left            =   4245
      TabIndex        =   98
      Top             =   855
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
      Object.Height          =   34
      Object.Width           =   84
   End
   Begin VB.Timer Timer2 
      Left            =   -30
      Top             =   825
   End
   Begin Cis3D_v60.CIS3D K_Sykbnm 
      Height          =   795
      Left            =   645
      Top             =   480
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   1402
      Caption         =   "読取"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   20.25
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
   Begin Cis3D_v60.CIS3D UC_3D4 
      Height          =   405
      Left            =   7980
      Top             =   -15
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   714
      BackColor       =   0
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
      cBoderWidth     =   3
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D H1lb_ADDCnt 
         Height          =   300
         Left            =   3030
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         ForeColor       =   12583104
         Caption         =   "ZZ9"
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
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   300
         Left            =   1845
         Top             =   60
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         ForeColor       =   12583104
         Caption         =   "手 入 力"
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
      Begin Cis3D_v60.CIS3D H1lb_OCRCnt 
         Height          =   300
         Left            =   1230
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         Caption         =   "ZZ9"
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
      Begin Cis3D_v60.CIS3D UC_3D5 
         Height          =   300
         Left            =   45
         Top             =   60
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         Caption         =   "QR読取"
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
   Begin VB.PictureBox B1_Area3 
      Height          =   1005
      Left            =   4575
      ScaleHeight     =   945
      ScaleWidth      =   1380
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   1440
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   840
         Index           =   8
         Left            =   60
         Top             =   60
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1482
         ForeColor       =   16711680
         Caption         =   "ページ"
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
         Begin CisText_V60.CisText iB1_PageS 
            Height          =   375
            Left            =   135
            TabIndex        =   90
            Top             =   360
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   661
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
            Text            =   "99"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin CisText_V60.CisText iB1_PageE 
            Height          =   375
            Left            =   750
            TabIndex        =   91
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
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
            Text            =   "99"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "/"
            Height          =   315
            Left            =   570
            TabIndex        =   97
            Top             =   465
            Width           =   105
         End
      End
   End
   Begin VB.PictureBox B2_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Left            =   4980
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1035
      Begin VB.TextBox iB2_Input 
         Height          =   330
         Left            =   60
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   75
         Width           =   825
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
      Left            =   8595
      ScaleHeight     =   660
      ScaleWidth      =   5190
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   480
      Width           =   5250
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2610
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
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
         Left            =   3465
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
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
         Left            =   4320
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
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
         Left            =   1755
         Top             =   30
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   885
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "手入力"
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
      Begin CisBtn_60.CisBtn PB_RED 
         Height          =   585
         Left            =   30
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "読取"
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
         PFBackColor     =   8388736
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
      TabIndex        =   93
      Text            =   "X"
      Top             =   450
      Width           =   345
   End
   Begin VB.Timer Timer1 
      Left            =   -30
      Top             =   405
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
      Caption         =   "【 受注納品書入力 [丸高]  】"
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
      Height          =   1920
      Left            =   630
      Top             =   1425
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   3387
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
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   870
         Left            =   2610
         Top             =   960
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   1535
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
      End
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   870
         Index           =   0
         Left            =   90
         Top             =   960
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1535
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_Ymd 
            Height          =   345
            Left            =   90
            TabIndex        =   8
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
         Begin CisText_V60.CisText iH1_Bin 
            Height          =   360
            Left            =   2055
            TabIndex        =   9
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
   Begin VB.Label ReadColor 
      BackColor       =   &H00FFC0C0&
      Caption         =   "読取中色"
      Height          =   255
      Left            =   5040
      TabIndex        =   99
      Top             =   570
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "CMA0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   受注納品書入力
'**       フォームID    :   CMA0010
'**       処理概要      :
'**
'**       作  成  日    :   2008/09/11  By CIS SIMPR-A
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKb         As String       ' 処理区分(追加･修正･削除)
    Dim SyoriNM         As String       ' 処理名称(追加･修正･削除)
    Dim SyoriBefore     As String

    Dim wEndChk         As Boolean
    Dim wIndex          As Integer
    Dim wJyuSu          As Long
    Dim wInput          As Boolean      ' 入力ﾌﾗｸﾞ

    Dim ErrCnt          As Long
    Dim wNum1           As Integer
    Dim wNum2           As Integer
    Dim wNumL           As Long
    Dim wStr1           As String
    Dim sv_Uke2         As String * 2   ' 受入(2桁)
    Dim sv_NOTor        As String       ' 納入先

    Dim wcc_Head        As String
    Dim wcc_Body        As String
    Dim wcc_Body1       As String       ' 2008.09.12 add : PXJ0040で10行のものを15行にするので追加
    Dim wcc_Body2       As String
    Dim wcce_Obj        As Object
    
    Dim mJKNo           As Long
    
    Dim wRowSu          As Integer      ' 2008.09.11 add : 行数
    Dim wMtFlg          As Boolean      ' 2008.10.06 add : True=名称マスタに変換先コードあり／False=なし
    
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

    wRowSu = 14     ' 2008.09.11 add : 行数(15行:0～14)
    
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
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = True    ' 処理区分点滅禁止
    
    B2_Area1.Left = -1000
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    Call QT10QSRInfoGet
    H1Op_JKbn(1).Value = True
    ProcHB = "H1"
    If Not QT10QSR_UseFlg Then
       SyoriKb = "ADD"
    Else
       SyoriKb = "RED"
    End If
    Call DispChange(ProcHB)
    Call SyoriName(SyoriKb)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

'   >> 名称マスタ確認
    Call MeisyoInit
'   >> 日付表示
'    iH1_Ymd = Format(Date, "yyyymmdd")
'    iH1_Bin = "01"
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct

    If InitProc Then        ' 初期処理（データ削除）
        wEndChk = True
    Else
        wEndChk = False
        Call PB_END_Click
    End If

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
        Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【追加】
        Case vbKeyF5:       If PB_RED.Visible Then Call PB_RED_Click    '【読取】
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
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
    Cancel = 1
    Call PB_END_Click
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'   #### 終了処理 ####
'    If wEndChk Then Call EndProc
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
    
    If ProcHB = "B1" Then
        If SyoriKb <> "RED" Then
           ProcHB = "H1"
           Call SyoriName(SyoriKb)
           Call DispChange(ProcHB)
        Else
           ProcHB = "B2"
           Call SyoriName(SyoriKb)
           Call DispChange(ProcHB)
        End If
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKb)
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    SyoriKb = "ADD"
    Call SyoriName(SyoriKb)
    If ProcHB = "B2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
    End If
End Sub
'+----------------------------+
'+      読取 キ ー(F5)         +
'+----------------------------+
Private Sub PB_RED_Click()
    SyoriKb = "RED"
    Call SyoriName(SyoriKb)
    If ProcHB = "B2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag Like "*iB12_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag Like "iB1_Seban*" Then Call Look_Hinbn: Exit Sub
'
    If PB_Look.Tag = "iH1_Ymd" Then iH1_Ymd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iB1_NSYmd" Then iB1_NSYmd.ShowCalender: Exit Sub
End Sub
'+----------------------------+
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()
    wIndex = CisFun.Val2(Mid(PB_Look.Tag, 1, 2))

    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 0
    RVI_SK = 9
    RVI_Torcd = ""
    RVI_Ukeir = ""
    RVI_Seban = ""
    RVI_Tehai = ""

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iB12_Seban(wIndex) = RTrim(RV_Seban)
       iB12_Hinbn(wIndex) = RTrim(RV_Hinbn)
    End If
End Sub
'+----------------------------------------+
'+  H1Op_Jkbn ( Click )      受注区分
'+----------------------------------------+
Private Sub H1Op_JKbn_Click(Index As Integer)
    H1Op_JKbn(1).Tag = Index
    For wIndex = 1 To 8
       H1Op_JKbn(wIndex).BackColor = &H8000000F
    Next wIndex
    H1Op_JKbn(Index).BackColor = &H8080FF
    Select Case Index
        Case 1:     Back_Ymd.Caption = "納 入 日"
        Case Else:  Back_Ymd.Caption = "納 入 日"
    End Select
End Sub
'+-----------------------------+
'+  iH1_Ymd - 受注日・納入日
'+-----------------------------+
Private Sub iH1_Ymd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ymd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub

'************************************'
'*        画面フィールド
'************************************'
'--------------------------------------------------------------- 2008.09.11 add start
'+-----------------------------+
'+  iB1_NTorcd - 納入先
'+-----------------------------+
Private Sub iB1_NTorcd_LostFocus()
    If iB1_NTorcd <> "" Then
        gSL_Select = "SELECT * FROM 名称マスタ " & _
                     " WHERE 区分名称 = '丸高変換取引先コード' " & _
                     "   AND 値       = '" & RTrim(iB1_NTorcd) & "'"
        If SYMRead(gSL_Select, 1) Then
           Blb_HCord = SYM.英数字1
        Else
           Blb_HCord = ""
        End If
    End If
End Sub
'--------------------------------------------------------------- 2008.09.11 add end

'+-----------------------------+
'+  iB1_NSYmd - 納入指示日
'+-----------------------------+
Private Sub iB1_NSYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_NSYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*---------------------------------------------------*
'*      iB12_Hinbn   ( GotFocus )        品番
'*---------------------------------------------------*
Private Sub iB12_Hinbn_GotFocus(Index As Integer)
    PB_Look.Tag = Format(Index, "00") & ActiveControl.Name
    PB_Look.Visible = True
End Sub
'*---------------------------------------------------*
'*      iB12_Hinbn   ( LostFocus )       品番
'*---------------------------------------------------*
Private Sub iB12_Hinbn_LostFocus(Index As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'*---------------------------------------------------*
'*      iB12_Hako   ( LostFocus )       箱数
'*---------------------------------------------------*
Private Sub iB12_Hako_LostFocus(Index As Integer)
    If RTrim(iB12_Hinbn(Index)) = "" Then Exit Sub
'   * 省略時 １ *
    If CisFun.Val2(iB12_Hako(Index)) = 0 Then
        If CisFun.Val2(iB12_Su(Index)) <> 0 And CisFun.Val2(iB12_Syuyo(Index)) <> 0 Then
            gLong = CisFun.Val2(iB12_Su(Index)) \ CisFun.Val2(iB12_Syuyo(Index))
            If CisFun.Val2(iB12_Su(Index)) Mod CisFun.Val2(iB12_Syuyo(Index)) <> 0 Then
                gLong = gLong + 1
            End If
            iB12_Hako(Index) = CisFun.RSetFld(gLong, 0, iB12_Hako(Index).cFormat)
        Else
           iB12_Hako(Index) = "  1"
        End If
    End If
'   * 納入数計算 *
    If CisFun.Val2(iB12_Su(Index)) <> 0 Then Exit Sub
    wNumL = CisFun.Val2(iB12_Hako(Index)) * CisFun.Val2(iB12_Syuyo(Index))
    iB12_Su(Index) = Format(Format(wNumL, "#"), String(7, "@"))
End Sub
'*---------------------------------------------------*
'*      iB12_Syuyo  ( LostFocus )       収容数
'*---------------------------------------------------*
Private Sub iB12_Syuyo_LostFocus(Index As Integer)
    If RTrim(iB12_Hinbn(Index)) = "" Then Exit Sub
'   * 省略時 １ *
    If CisFun.Val2(iB12_Syuyo(Index)) = 0 Then
       If CisFun.Val2(iB12_Su(Index)) <> 0 And CisFun.Val2(iB12_Hako(Index)) <> 0 Then
          gLong = CisFun.Val2(iB12_Su(Index)) / CisFun.Val2(iB12_Hako(Index))
          iB12_Syuyo(Index) = CisFun.RSetFld(gLong, 0, iB12_Syuyo(Index).cFormat)
       Else
          iB12_Syuyo(Index) = Format(Format(CisFun.Val2(iB12_Su(Index)), "#"), String(5, "@"))
       End If
    End If
'   * 納入数計算 *
    If CisFun.Val2(iB12_Su(Index)) <> 0 Then Exit Sub
    wNumL = CisFun.Val2(iB12_Hako(Index)) * CisFun.Val2(iB12_Syuyo(Index))
    iB12_Su(Index) = Format(Format(wNumL, "#"), String(7, "@"))
End Sub
'*---------------------------------------------------*
'*      iB12_Su     ( LostFocus )       納入数
'*---------------------------------------------------*
Private Sub iB12_Su_LostFocus(Index As Integer)
    If CisFun.Val2(iB12_Su(Index)) <> 0 Then Exit Sub

    wNumL = CisFun.Val2(iB12_Hako(Index)) * CisFun.Val2(iB12_Syuyo(Index))
    iB12_Su(Index) = Format(Format(wNumL, "#"), String(7, "@"))
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
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   
    Dim W_Color     As Long
    
    W_Color = &H80000008 '初期色ｾｯﾄ(黒)
    K_Sykbnm.BackColor = &HC0C0C0
    Select Case SyoriKb
     Case "RED"
          SyoriNM = "QR読取"
          K_Sykbnm.ForeColor = W_Color   '(黒)
          PB_ADD.Visible = True
          PB_RED.Visible = False
     Case "ADD"
          SyoriNM = "手入力"
          K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
          PB_ADD.Visible = False
          PB_RED.Visible = True
     Case Else
          SyoriNM = ""
    End Select
    
    Timer2.Enabled = False
    If SyoriNM = "" Then
        K_Sykbnm.Visible = False
        K_Sykbnm.Caption = ""
        Exit Sub
    End If
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    Timer2.Enabled = True
    
    If Not QT10QSR_UseFlg Then
       PB_ADD.Visible = False
       PB_RED.Visible = False
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

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
    B2Mode = False: B2Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1":   H1Mode = True:  H1Color = gPGotSel
       Case "B1":   B1Mode = True:  B1Color = gPGotSel
       Case "B2":   B2Mode = True:  B2Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode
    B1_Area2.BackColor = B1Color
    B1_Area3.Enabled = B1Mode
    B1_Area3.BackColor = B1Color
    B2_Area1.Enabled = B2Mode
    B2_Area1.BackColor = B2Color

    PB_CAN.Visible = Not (H1Mode)

    If ProcHB = "B2" Then
        SyoriNM = "【 読 取 中 】"
        K_Sykbnm.BackColor = ReadColor.BackColor
        Set QT10QSR_Comm = PDQComm1
        Call QT10QSRInfoGet
        If Not QT10QSR_Open Then
           End
        End If
    End If
    
    PB_ADD.Visible = H1Mode And Not PB_RED.Visible  ' 2008.09.11 upd : And～追加
    PB_RED.Visible = H1Mode And Not PB_ADD.Visible  ' 2008.09.11 upd : And～追加
    
    PB_ENT.Visible = Not B2Mode
    
    If Not QT10QSR_UseFlg Then
       PB_ADD.Visible = False
       PB_RED.Visible = False
    End If
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
        iH1_Ymd = Format(Date, "yyyymmdd")
        iH1_Bin = "01"
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
'   * 受注区分 *
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        If SyoriKb = "RED" Then
            ProcHB = "B2"
            Call DispChange(ProcHB)
            Call ChiketoProc                         ' QR読取処理
            GoTo ReturnPress_Ed
        End If
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If

'   * 伝票入力 *
    If ProcHB = "B1" Then
        Call Line_Cut
        If Not ContentsChk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        
'--------------------------------------------------------------- 2008.10.06 add start
        '丸高変換取引先コードの作成（納入先に対する該当がない場合）
        If wMtFlg = False Then
           gSL_Select = "SELECT * FROM 名称マスタ " & _
                        " WHERE 区分名称   = '丸高変換取引先コード' " & _
                        "   AND 区分タイプ = 'K' "
            
            If Not SYMRead(gSL_Select, 1) Then
               ' 丸高変換取引先コードの項目自体が名称マスタに登録されていないので作成する
               Call ItemsClearSYMCreate
               With SYMCreate
                   .区分名称 = "丸高変換取引先コード"
                   .区分タイプ = "K"
                   .区分桁数 = 7
                   .作成者 = gTanto
                   .英数字1 = ""
                   If Not SYMInitCreate(True) Then
                   End If
               End With
            End If
            
            ' 丸高変換取引先コードが名称マスタに登録されていないので作成する
            Call ItemsClearSYMCreate
            With SYMCreate
                .区分名称 = "丸高変換取引先コード"
                .区分桁数 = 7
                .値 = iB1_NTorcd
                .値名称 = ""
                .作成者 = gTanto
                .検索名 = ""
                .コード名 = ""
                .名称名 = ""
                .英数字1 = ""
                If Not SYMInitCreate() Then     'If Not SYMInitCreate(True) Then
                End If
            End With
        End If
'--------------------------------------------------------------- 2008.10.06 add end
        
        Call HeadBodyClear("B")
        If ErrCnt = 0 And (SyoriKb = "MNT" Or SyoriKb = "DEL") Then
            SyoriKb = "RED"
            ProcHB = "H1"
        Else
            If SyoriKb = "ADD" Then ProcHB = "H1" Else ProcHB = "B2"
        End If
        
        Call SyoriName(SyoriKb)
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

    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With

'    If RTrim(iH1_Ymd) = "" Then iH1_Ymd = Format(Now(), "yyyymmdd")
'    If CisFun.Val2(iH1_Bin) = 0 Then iH1_Bin = "01"
'   << 受注日 or 納入日 >>
    If RTrim(iH1_Ymd) = "" And RTrim(iH1_Bin) = "" Then
        Head1Chk = True
        Exit Function
    End If
    If RTrim(iH1_Ymd) = "" Then
        CisFun.MB_MSG(2) = "    " & Replace(Back_Ymd.Caption, " ", "") & "を入力して下さい。      "
        CisFun.MBOX
        iH1_Ymd.SetFocus
        Exit Function
    Else
        If Not iH1_Ymd.cChkResult Then
            CisFun.MB_MSG(2) = "    " & Replace(Back_Ymd.Caption, " ", "") & "を正しく入力して下さい            "
            CisFun.MBOX
            iH1_Ymd.SetFocus
            Exit Function
        End If
    End If
'   << 便 >>
    If Trim(iH1_Bin) = "" Then
        CisFun.MB_MSG(2) = "    便を入力して下さい            "
        CisFun.MBOX
        iH1_Bin.SetFocus
        Exit Function
    End If

    With CisFun
        gStr = ""
        .MB_Lines = 5
        Select Case Kdou_Check(iH1_Ymd)
            Case 1: gStr = "      指定日付は稼働日ではありません          "
            Case 2: gStr = "      カレンダマスタ未登録                "
        End Select
    
        If Trim(gStr) <> "" Then
            .MB_MSG(2) = gStr
            .MB_MSG(4) = "      よろしいですか？            "
            .MB_Button = OK_CAN
            .MB_Title = ""
            
            If Not .MBOX Then
               iH1_Ymd.SetFocus
               Exit Function
            End If
        End If
    End With

    Head1Chk = True
End Function
'+----------------------------------------------------------------------+
'+          明 細 行  行 詰 め
'+----------------------------------------------------------------------+
Private Sub Line_Cut()
    Static wlc_Flg      As Boolean
    
    For wNum1 = 0 To wRowSu                     ' 2008.09.11 upd : 9 → wRowSu
        If RTrim(iB12_Hinbn(wNum1)) = "" And _
           RTrim(iB12_Seban(wNum1)) = "" And CisFun.Val2(iB12_Hako(wNum1)) = 0 And CisFun.Val2(iB12_Syuyo(wNum1)) = 0 And _
           CisFun.Val2(iB12_Su(wNum1)) = 0 Then

            wlc_Flg = False
            For wNum2 = wNum1 To wRowSu
                If RTrim(iB12_Hinbn(wNum2)) <> "" Or _
                   RTrim(iB12_Seban(wNum2)) <> "" Or CisFun.Val2(iB12_Hako(wNum2)) <> 0 Or CisFun.Val2(iB12_Syuyo(wNum2)) Or _
                   CisFun.Val2(iB12_Su(wNum2)) <> 0 Then

                    iB12_Hinbn(wNum1) = RTrim(iB12_Hinbn(wNum2))
                    iB12_Seban(wNum1) = RTrim(iB12_Seban(wNum2))
                    iB12_Hako(wNum1) = iB12_Hako(wNum2)
                    iB12_Syuyo(wNum1) = iB12_Syuyo(wNum2)
                    iB12_Su(wNum1) = iB12_Su(wNum2)

                    iB12_Hinbn(wNum2) = ""
                    iB12_Seban(wNum2) = ""
                    iB12_Hako(wNum2) = "":      iB12_Syuyo(wNum2) = ""
                    iB12_Su(wNum2) = ""

                    wlc_Flg = True
                    Exit For
                End If
            Next wNum2
            If Not wlc_Flg Then Exit For
        End If
    Next wNum1
End Sub
'*----------------------------------------------------------------*
'*          伝票内容チェック
'*
'*      画面表示項目をチケット読取データの様に、文字列にして
'*                          ストアドにて内容のチェックを行う。
'*      チェック後に、ストアドからの Outputﾊﾟﾗﾒｰﾀを使用し、
'*                                              再表示をする。
'*----------------------------------------------------------------*
Private Function ContentsChk() As Boolean
    ContentsChk = False
 
    Dim wErNo           As Integer
    Dim wErGyo          As Integer
    Dim BodySetFlg      As Boolean
    Dim wcc_Body1       As String       ' 2008.09.11 add
    Dim wk_Head         As String       ' 2008.09.12 add
    Dim wk_Body         As String       ' 2008.09.12 add
    
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
    
    If RTrim(iB1_Noubn) = "" Then
        CisFun.MB_MSG(2) = "  納入番号を入力して下さい      "
        CisFun.MBOX
        iB1_Noubn.SetFocus
        Exit Function
    End If
    If RTrim(iB1_NTorcd) = "" Then
        CisFun.MB_MSG(2) = "  納入先を入力して下さい      "
        CisFun.MBOX
        iB1_NTorcd.SetFocus
        Exit Function
'--------------------------------------------------------------- 2008.09.11 add start
    Else
        gSL_Select = "SELECT * FROM 名称マスタ " & _
                     " WHERE 区分名称 = '丸高変換取引先コード' " & _
                     "   AND 値       = '" & RTrim(iB1_NTorcd) & "'"
        
        wMtFlg = True                   ' 名称マスタに丸高変換取引先コードがあるか確認
        If SYMRead(gSL_Select, 1) Then
            Blb_HCord = SYM.英数字1
        Else
            wMtFlg = False
            Blb_HCord = ""
'            ' 名称マスタに登録されていないので作成する -------- 2008.10.06 del start
'            Call ItemsClearSYMCreate
'            With SYMCreate
'                .区分名称 = "丸高変換取引先コード"
'                .区分桁数 = 7
'                .値 = iB1_NTorcd
'                .値名称 = ""
'                .作成者 = gTanto
'                .検索名 = ""
'                .コード名 = ""
'                .名称名 = ""
'                .英数字1 = "1234567"
'                If Not SYMInitCreate(True) Then
'                End If
'            End With   ---------------------------------------- 2008.10.06 del end
        End If
'--------------------------------------------------------------- 2008.09.11 add end
    End If
    
    If RTrim(iB1_NSYmd) = "" Then
        CisFun.MB_MSG(2) = "  納入指示日を入力して下さい      "
        CisFun.MBOX
        iB1_NSYmd.SetFocus
        Exit Function
    End If
    If RTrim(iB1_NSYmd) <> "" And Not iB1_NSYmd.cChkResult Then
        CisFun.MB_MSG(2) = "  納入指示日エラー      "
        CisFun.MBOX
        iB1_NSYmd.SetFocus
        Exit Function
    End If
    If RTrim(iB1_JTorcd) = "" Then
        CisFun.MB_MSG(2) = "  受注者を入力して下さい      "
        CisFun.MBOX
        iB1_JTorcd.SetFocus
        Exit Function
    End If
    
'*--------------------------------------------------<< 2004/06/28 Insert Start >>
    If CisFun.Val2(iB1_PageS) = 0 Then iB1_PageS = 1
    If CisFun.Val2(iB1_PageE) = 0 Then iB1_PageE = 1
'*--------------------------------------------------<< 2004/06/28 Insert End >>
'
'===* ボディ部作成 *===================================
'
   '一旦QRデータと同じ物を作成
    wcc_Body = "": wcc_Body2 = "": BodySetFlg = False
    iB1_PageS.Tag = ""
    For wNum1 = 0 To wRowSu
'       * 品番 *
        If RTrim(iB12_Hinbn(wNum1)) = "" Then
            wcc_Body = wcc_Body & Space(18)
        Else
            wcc_Body = wcc_Body & Format(RTrim(iB12_Hinbn(wNum1)), "!" & String(12, "@")) & Space(6)
            BodySetFlg = True
            ' 明細行数
            iB1_PageS.Tag = wNum1 + 1
        End If
'       * 背番号*
        If RTrim(iB12_Seban(wNum1)) = "" Then
            wcc_Body = wcc_Body & Space(4)
        Else
            wcc_Body = wcc_Body & Format(RTrim(iB12_Seban(wNum1)), "!@@@@")
            BodySetFlg = True
        End If
'       * 収容数 *
        wcc_Body = wcc_Body & Format(CisFun.Val2(iB12_Syuyo(wNum1)), "00000")

'       * 箱数 *
        wcc_Body = wcc_Body & Format(CisFun.Val2(iB12_Hako(wNum1)), "000")

'       * 納入数 *
        wcc_Body = wcc_Body & Format(CisFun.Val2(iB12_Su(wNum1)), "0000000")
    Next wNum1
'Call Test_LogOutput(wcc_Body + "$$$$", 0)
    
    If Not BodySetFlg Then
        With CisFun
             .MB_Lines = 5
             .MB_MSG(2) = "  明細行を入力して下さい      "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
        End With
        iB12_Hinbn(0).SetFocus
        Exit Function
    End If
    
'===* ヘッド部作成 *===================================
'
   '一旦QRデータと同じ形を作成
    wk_Head = ""
    If Trim(iB1_Noubn) = "" Then    '------------------------------- 納入番号(5)
       wk_Head = wk_Head & Space(5)
    Else
       wk_Head = wk_Head & Format(RTrim(iB1_Noubn), "!@@@@@")
    End If
    If Trim(iB1_NSYmd) = "" Then    '------------------------------- 納入指示日(8)
       wk_Head = wk_Head & Space(8)
    Else
       wk_Head = wk_Head & Format(RTrim(iB1_NSYmd), "!@@@@@@@@")
    End If
    wk_Head = wk_Head & "0    "     '------------------------------- ？？(5)
    If Trim(iB1_NTorcd) = "" Then   '------------------------------- 納入先(5)
       wk_Head = wk_Head & Space(5)
    Else
      'wk_Head = wk_Head & Format(RTrim(iB1_NTorcd), "!@@@@")
      'wk_Head = wk_Head & Format(RTrim(iB1_NTorcd), "@@@@@")
       wk_Head = wk_Head & Mid(RTrim(iB1_NTorcd) + Space(5), 1, 5)
    End If
    If Trim(iB1_JTorcd) = "" Then   '------------------------------- 受注者(5)
       wk_Head = wk_Head & Space(5)
    Else
      'wk_Head = wk_Head & Format(RTrim(iB1_JTorcd), "!@@@@@")
      'wk_Head = wk_Head & Format(RTrim(iB1_JTorcd), "@@@@@")
       wk_Head = wk_Head & Mid(RTrim(iB1_JTorcd) + Space(5), 1, 5)
    End If
    wk_Head = wk_Head & "01   "     '------------------------------- ？？(5)
    If Trim(iB1_NBin) = "" Then     '------------------------------- 納入指示便(2)
       wk_Head = wk_Head & Space(2)
    Else
       wk_Head = wk_Head & Format(RTrim(iB1_NBin), "00")
    End If
    wk_Head = wk_Head & "   "       '------------------------------- ？？(3)

   'ストアドに渡すヘッダ形式を作成
    Dim wk As String
    wcc_Head = ""                                   ' 項目　　　実桁→渡桁
   'wk = Trim(Mid(wk_Head, 19, 5))
    wcc_Head = wcc_Head & Mid(wk_Head, 19, 5)       ' 納入先　　(       5)
    wcc_Head = wcc_Head & Space(3)                  ' 納入事業所(       3)
    wcc_Head = wcc_Head & Mid(wk_Head, 6, 8)        ' 納入指示日(  8→  8)
    wcc_Head = wcc_Head & Mid(wk_Head, 34, 2)       ' 納入便　　(  2→  2)
    wcc_Head = wcc_Head & Space(3)                  ' 機番　　　(       3)
    wcc_Head = wcc_Head & Space(1)                  ' ＃　　　　(       1)
    wcc_Head = wcc_Head & Space(4)                  ' 支給元　　(       4)
   'wk = Trim(Mid(wk_Head, 24, 5))
    wcc_Head = wcc_Head & Mid(wk_Head, 24, 5)       ' 受注者　　(       5)
   'wcc_Head = wcc_Head & Space(1)                  ' 受事　　　(       1)
    wcc_Head = wcc_Head & Mid(wk_Head, 1, 5)        ' 納入番号　(       5)
    wcc_Head = wcc_Head & "01"                      ' ページ S  (       2)
    wcc_Head = wcc_Head & "01"                      ' ページ E  (       2)
    wcc_Head = wcc_Head & _
        Format(CisFun.Val2(iB1_PageS.Tag), "00")    ' 入力行数  (       2)
    wcc_Head = wcc_Head & "0"                       ' 発注区分  (       1)
'Call Test_LogOutput(wk_Head + vbCrLf + wcc_Head + "$$$$", 0)
'
    If ChekData(1) Then
       ContentsChk = True
    End If
End Function
'+------------------------------------------+
'+
'+------------------------------------------+
Private Function ChekData(ReadType As Integer) As Boolean
    Dim wErNo           As Integer
    Dim wErGyo          As Integer
                
    ChekData = False
    
    gStr = wcc_Body
    wcc_Body = Mid(gStr, 1, 185)    ' 37 * 5
    wcc_Body1 = Mid(gStr, 186, 185) ' 2008.09.12 add
    wcc_Body2 = Mid(gStr, 371)
'Call Test_LogOutput(wcc_Head + vbCrLf + wcc_Body + vbCrLf + wcc_Body1 + vbCrLf + wcc_Body2, 0)

    sv_Uke2 = ""
    sv_NOTor = ""

    With CisDB
        .SQL = "受注丸高伝票チェック"
        .StoadoCount = 12               ' 2008.09.11 11→12
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用      ' 処理区分 0:ﾘｰﾀﾞｰ,1:手入力
        .ParaNo = 2: .ParaIO = Input用      ' 受注区分
        .ParaNo = 3: .ParaIO = Input用      ' 納入日
        .ParaNo = 4: .ParaIO = Input用      ' 納入便
        .ParaNo = 5: .ParaIO = Input用      ' 納入先～受注者事業所・納品書№・ページ
        .ParaNo = 6: .ParaIO = InOut用      ' 伝票の明細部(品番～納入数<5明細分>)
        .ParaNo = 7: .ParaIO = InOut用      ' 伝票の明細部(品番～納入数<5明細分>)
        .ParaNo = 8: .ParaIO = InOut用      ' 伝票の明細部(品番～納入数<5明細分>)   ' 2008.09.11 add : 以降、番号+1した
        .ParaNo = 9: .ParaIO = OutPut用     ' 変換後 受入
        .ParaNo = 10: .ParaIO = OutPut用     ' 変換後 納入先
        .ParaNo = 11: .ParaIO = OutPut用     ' チェックエラーNO
        .ParaNo = 12: .ParaIO = OutPut用     ' チェックエラー行NO
        
        .ParaNo = 1: .ParaValue = ReadType   ' 0:ﾘｰﾀﾞｰ,1:手入力
        .ParaNo = 2: .ParaValue = H1Op_JKbn(1).Tag
        .ParaNo = 3: .ParaValue = iH1_Ymd
        .ParaNo = 4: .ParaValue = iH1_Bin
        .ParaNo = 5: .ParaValue = wcc_Head
        .ParaNo = 6: .ParaValue = wcc_Body
        .ParaNo = 7: .ParaValue = wcc_Body1
        .ParaNo = 8: .ParaValue = wcc_Body2

        If Not .DBStored Then End
        
        ' 受入
        .ParaNo = 9: sv_Uke2 = RTrim(.ParaValue)
        ' 納入先
        .ParaNo = 10: sv_NOTor = RTrim(.ParaValue)
        ' 明細1-5
        .ParaNo = 6: wStr1 = Mid(.ParaValue, 1, 185)
        ' 明細6-10
        .ParaNo = 7: wStr1 = wStr1 & Mid(.ParaValue, 1, 185)
        ' 明細11-15
        .ParaNo = 8: wStr1 = wStr1 & .ParaValue
        ' エラー番号
        .ParaNo = 11: wErNo = CisFun.Val2(.ParaValue)
        ' エラー行No
        .ParaNo = 12: wErGyo = CisFun.Val2(.ParaValue)
'Call Test_LogOutput(wStr1, 0)
    End With

'   * 画面再表示 *
    Call HeadBodyClear("B")
'
    iB1_NTorcd = Mid(wcc_Head, 1, 5)
    iB1_NSYmd = Mid(wcc_Head, 9, 8)
    iB1_NBin = Mid(wcc_Head, 17, 2)
    iB1_JTorcd = Mid(wcc_Head, 27, 5)
    iB1_Noubn = Mid(wcc_Head, 32, 5)
    iB1_PageS = CisFun.RSetFld(Mid(wcc_Head, 37, 2), 0, iB1_PageS.cFormat)
    iB1_PageE = CisFun.RSetFld(Mid(wcc_Head, 39, 2), 0, iB1_PageE.cFormat)
    iB1_PageS.Tag = CisFun.Val2(Mid(wcc_Head, 41, 2))    ' 行数
'   iB1_NUkeir = Mid(wcc_Head, 5, 4)    ' 2008.09.11 del
'   iB1_Kibn = Mid(wcc_Head, 19, 3)     ' 2008.09.11 del
'   iB1_Sharp = Mid(wcc_Head, 22, 1)    ' 2008.09.11 del
'   iB1_STorcd = Mid(wcc_Head, 23, 4)   ' 2008.09.11 del
'   iB1_JKojyo = Mid(wcc_Head, 31, 1)   ' 2008.09.11 del
    Call iB1_NTorcd_LostFocus           ' 2008.10.06 add : 納入先→変換先コード　再検索＆表示
    
    For wNum1 = 0 To wRowSu
        wNum2 = wNum1 * 37
        If RTrim(Mid(wStr1, wNum2 + 1, 22)) = "" Then Exit For
        iB12_Hinbn(wNum1) = RTrim(Mid(wStr1, wNum2 + 1, 18))
        iB12_Seban(wNum1) = RTrim(Mid(wStr1, wNum2 + 19, 4))
        iB12_Syuyo(wNum1) = CisFun.RSetFld(Mid(wStr1, wNum2 + 23, 5), 0, iB12_Syuyo(wNum1).cFormat)
        iB12_Hako(wNum1) = CisFun.RSetFld(Mid(wStr1, wNum2 + 28, 3), 0, iB12_Hako(wNum1).cFormat)
        iB12_Su(wNum1) = CisFun.RSetFld(Mid(wStr1, wNum2 + 31, 7), 0, iB12_Su(wNum1).cFormat)
    Next wNum1
    
    If wErNo = 0 Then
        Static HasuChk      As Integer
        
        HasuChk = 90
        For wNum1 = 0 To wRowSu
            If CisFun.Val2(iB12_Syuyo(wNum1)) <> 0 Then
'               * 端数･箱数チェック*
                wNumL = CisFun.Val2(iB12_Su(wNum1)) \ CisFun.Val2(iB12_Syuyo(wNum1))
                If (CisFun.Val2(iB12_Su(wNum1)) Mod CisFun.Val2(iB12_Syuyo(wNum1))) Then
                    wNumL = wNumL + 1
                End If
            
                If CisFun.Val2(iB12_Hako(wNum1)) <> wNumL Then
                    With CisFun
                        .MB_Lines = 5
                        .MB_MSG(1) = "【  " & StrConv(wNum1 + 1, vbWide) & " 行目  】"
                        .MB_MSG(3) = "            箱数エラー "
                        .MBOX
                    End With
                    iB12_Hako(wNum1).SetFocus
                    Exit Function
                End If
            End If
            
'           * 端数確認メッセージ *
            wNumL = CisFun.Val2(iB12_Syuyo(wNum1)) * CisFun.Val2(iB12_Hako(wNum1))
            If CisFun.Val2(iB12_Su(wNum1)) <> wNumL Then
                If HasuChk = 90 Then HasuChk = wNum1
            End If
        Next wNum1
        
        If HasuChk <> 90 Then
            With CisFun
                 .MB_Lines = 5
                 .MB_MSG(1) = "【収容数 × 箱数 ＝ 納入数】"
                 .MB_MSG(2) = "          になっていない行が有ります   "
                 .MB_MSG(4) = "         処理を続行しますか？"
                 .MB_Button = OK_CAN
                 If Not .MBOX Then iB12_Hako(HasuChk).SetFocus: Exit Function
            End With
        End If
        ChekData = True
        Exit Function
    End If
    
    wNum1 = wErGyo
    wNum2 = wErGyo - 1
    
    With CisFun
        .MB_Lines = 5
        .MB_Title = ""
        .MB_Button = Error
    
        Select Case wErNo
            Case 1
                   '.MB_MSG(1) = "     納入先変換エラー"
                   '.MB_MSG(2) = "   ( 取引先受入マスタ )    "
                    .MB_MSG(1) = "     納入先未登録"
                    .MB_MSG(2) = "   ( 取引先マスタ )    "
                    Set wcce_Obj = iB1_NTorcd
            Case 2
                    .MB_MSG(1) = "     納入先未登録"
                    .MB_MSG(2) = "   ( 取引先マスタ )    "
                    Set wcce_Obj = iB1_NTorcd
            Case 3
                    .MB_MSG(2) = "   自社向け伝票でない    "
                    Set wcce_Obj = iB1_JTorcd
'--------------------------------------------------------------- 2008.09.11 del start
'            Case 4
'                    .MB_MSG(1) = "       受入未登録"
'                    .MB_MSG(2) = "   ( 取引先受入マスタ )    "
'                    Set wcce_Obj = iB1_NUkeir
'--------------------------------------------------------------- 2008.09.11 del end
            Case 7
                    .MB_MSG(1) = "【  " & StrConv(wNum1, vbWide) & " 行目  】"
                    .MB_MSG(3) = "   収容数が異なる    "
                    Set wcce_Obj = iB12_Syuyo(wNum2)
            Case 8
                    .MB_MSG(1) = "【  " & StrConv(wNum1, vbWide) & " 行目  】"
                    .MB_MSG(3) = "   納入数が異なる ( 箱数×収容数 )    "
                    Set wcce_Obj = iB12_Hako(wNum2)
            Case 15
                    .MB_MSG(1) = "【  " & StrConv(wNum1, vbWide) & " 行目  】"
                    .MB_MSG(3) = "       品番未登録 ( 品番取引先マスタ )     "
                    Set wcce_Obj = iB12_Hinbn(wNum2)
            Case 16
                    .MB_MSG(2) = "       既に受注データが存在します     "
                    Set wcce_Obj = iB1_Noubn
            Case 17
                    .MB_MSG(2) = "       伝票入力済みです      "
                    Set wcce_Obj = iB1_Noubn
        End Select
        .MBOX
    End With
    
    On Error Resume Next
    wcce_Obj.SetFocus
    Set wcce_Obj = Nothing
    On Error GoTo 0

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

'  ( 確認メッセージ )
    If Not DBCheckMsg Then Exit Function
    
    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    With CisDB
    '   <><><> 受注管理NOの獲得 <><><>      '2004/04/15 Add
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
            .ParaNo = 2: mJKNo = .ParaValue
        End If
    End With
    
    For wIndex = 0 To wRowSu
        If RTrim(iB12_Hinbn(wIndex)) <> "" Or _
           RTrim(iB12_Seban(wIndex)) <> "" Then
        '   >> DBｾｯﾄ
            If Not DBSet_JNW(wIndex) Then GoTo DBPut_Err
        '   >> 出力
            Call JNWInsert
        End If
    Next wIndex

    If SyoriKb = "RED" Then
       H1lb_OCRCnt = Format(CisFun.Val2(H1lb_OCRCnt) + 1, "#,###")
    Else
       H1lb_ADDCnt = Format(CisFun.Val2(H1lb_ADDCnt) + 1, "#,###")
    End If

    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    On Error GoTo 0

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------------+
'+    処理確認メッセージ（DBUpDate)
'+------------------------------------+
Private Function DBCheckMsg()
    DBCheckMsg = False
    
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(1) = ""
        .MB_MSG(3) = "    【 " & SyoriNM & "処理確認 】      "
        .MB_Title = SyoriNM & "処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    DBCheckMsg = True
End Function
'+------------------------------------+
'+      セット - 受注入力ワーク
'+------------------------------------+
Private Function DBSet_JNW(Index As Integer)
    DBSet_JNW = False

    Call ItemsClearJNW
'    With CisDB
'    '   <><><> 受注管理NOの獲得 <><><>    ' 2004/04/15 Del
'        .SQL = "番号獲得"
'        .StoadoCount = 3
'        .ParaNo = 0: .ParaIO = Return用
'        .ParaNo = 1: .ParaIO = Input用
'        .ParaNo = 2: .ParaIO = OutPut用
'        .ParaNo = 3: .ParaIO = Return用
'    '
'        .ParaNo = 1: .ParaValue = 11
'    '
'        If Not .DBStored Then
'            CisFun.MB_Lines = 4
'            CisFun.MB_MSG(2) = "    受注管理NOの獲得に失敗しました            "
'            CisFun.MB_Title = "採番処理"
'            CisFun.MB_Button = Error
'            CisFun.MBOX
'            Exit Function
'        Else
'            .ParaNo = 2: JNW.受注管理NO = .ParaValue
'        End If
'    End With
    With JNW
        .受注管理NO = mJKNo                 ' 2004/04/15 Add
        .受注管理行NO = Index + 1           ' 2004/04/15 Add
        .端末番号 = "' + substring(host_name(),1,20) + '"
        .受注区分 = CisFun.Val2(H1Op_JKbn(1).Tag)
        .受注日 = Format(Date, "YYYYMMDD")
        .受注便 = ""
        .納入日 = iB1_NSYmd
        .納入便 = Format(CisFun.Val2(iB1_NBin), "00")
        .納入先 = RTrim(sv_NOTor)
        .出荷予定日 = .納入日
        .出荷予定便 = .納入便
        .受入 = RTrim(sv_Uke2)
'       .機番 = RTrim(Mid(iB1_Kibn, 1, 2))  ' 2008.09.15 del
        .納入番号 = RTrim(iB1_Noubn)
        .納品書頁 = CisFun.Val2(iB1_PageS)
        .納品書行 = Index + 1
        .品番 = RTrim(iB12_Hinbn(Index))
        .背番号 = RTrim(iB12_Seban(Index))
        .枚数 = CisFun.Val2(iB12_Hako(Index))
        .収容数 = CisFun.Val2(iB12_Syuyo(Index))
        .納入数 = CisFun.Val2(iB12_Su(Index))
        If .収容数 <> 0 Then
           If .納入数 Mod .収容数 <> 0 Then
              gLong = .納入数 \ .収容数
              .端数 = .納入数 - .収容数 * gLong
           End If
        End If
        If .端数 <> 0 Then .枚数 = .枚数 + 1
        Select Case .受注区分
            Case 2, 4, 5, 6, 7, 8
                .エフ発行枚数 = .枚数
                .エフ発行端数 = .端数
            Case Else
                .エフ発行枚数 = 0
                .エフ発行端数 = 0
        End Select
        .作成種別 = "J"
'*---------------------------------------------<< 2004/04/17 Insert Start >>
        .納品書発行区分 = 9     '>対象外
'*---------------------------------------------<< 2004/04/17 Insert End >>
    End With

    DBSet_JNW = True
End Function

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
            .MB_MSG(4) = "      『受注納品書入力（号口）』処理を終了しますか？"
            .MB_MSG(5) = "      いいえの場合は、未累積データを削除し続行します"
            .MB_Title = "未累積データの確認"
            .MB_Button = OK_CAN
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
'**************************************
'*                                    *
'*      かんばんＪＡＭＡ読取処理        *
'*                                    *
'**************************************
Private Function ChiketoProc()
'   * かんばんＪＡＭＡ読取画面表示 *
    
' 232C ｵﾌﾞｼﾞｪｸﾄ初期設定
    Set QT10QSR_Comm = PDQComm1
    Call QT10QSRInfoGet
    If Not QT10QSR_Open Then
        Call PB_CAN_Click
    End If
    
    On Error GoTo 0
    Exit Function

OcrProc_Err:
    Call CisFun.ErrorBox("JAMA QR読取")
    End
End Function
'+------------------------------+
'+  読取処理     *
'+------------------------------+
Private Sub PDQComm1_OnComm()
    Static wPNo     As Integer
    Static Rtn      As String
    Dim wComm       As String
    
    gStr = QT10QSR_Event
    If gStr = "ERR" Then Exit Sub
    If gStr <> "RCV" Then Exit Sub
    
    If Not QT10QSR_RCVProc Then Exit Sub    ' CR が存在しない時抜ける
'Call Test_LogOutput(QT10QSR_ReadData, 1)

    Call QRConvert_MT(QT10QSR_ReadData)      ' 2008.09.12 upd: Call QRConvert(QT10QSR_ReadData)
    QT10QSR_ReadData = ""
    If Not QRProc Then Exit Sub
End Sub
'+------------------------------------------+
'+      QR処理(内容ﾁｪｯｸ & DB更新)
'+------------------------------------------+
Public Function QRProc() As Boolean
    
    QRProc = False
'   * ヘッド部作成 *
    With typQRAfter2
        If Trim(.納入先) = "" Then
           wcc_Head = Space(5)
        Else
           wcc_Head = Format(RTrim(.納入先), "!@@@@@")
        End If
        If Trim(.納入事業所) = "" Then
           wcc_Head = wcc_Head & Space(3)
        Else
           wcc_Head = wcc_Head & Format(RTrim(.納入事業所), "!@@@")
        End If
        If Trim(.納入指示日) = "" Then
            wcc_Head = wcc_Head & Space(8)
        Else
            wcc_Head = wcc_Head & Format(RTrim(.納入指示日), "!@@@@@@@@")
        End If
        If Trim(.納入便) = "" Then
            wcc_Head = wcc_Head & Space(2)
        Else
            wcc_Head = wcc_Head & Format(RTrim(.納入便), "!@@")
        End If
        If Trim(.機番) = "" Then
            wcc_Head = wcc_Head & Space(3)
        Else
            wcc_Head = wcc_Head & Format(RTrim(.機番), "!@@@")
        End If
        If Trim(.カード丼) = "" Then
            wcc_Head = wcc_Head & Space(1)
        Else
            wcc_Head = wcc_Head & Format(RTrim(.カード丼), "@")
        End If
        If Trim(.支給元) = "" Then
            wcc_Head = wcc_Head & Space(4)
        Else
            wcc_Head = wcc_Head & Format(RTrim(.支給元), "!@@@@")
        End If
        If Trim(.受注者) = "" Then
            wcc_Head = wcc_Head & Space(5)
        Else
            wcc_Head = wcc_Head & Format(RTrim(.受注者), "!@@@@@")
        End If
'        If Trim(.受事) = "" Then
'            wcc_Head = wcc_Head & Space(1)
'        Else
'            wcc_Head = wcc_Head & Format(RTrim(.受事), "@")
'        End If
        If Trim(.納入番号) = "" Then
            wcc_Head = wcc_Head & Space(5)
        Else
            wcc_Head = wcc_Head & Format(RTrim(.納入番号), "!@@@@@")
        End If
        wcc_Head = wcc_Head & Format(CisFun.Val2(.ページNO), "00")
        wcc_Head = wcc_Head & Format(CisFun.Val2(.総ページ数), "00")
        wcc_Head = wcc_Head & Format(CisFun.Val2(.明細行数), "00")     ' 明細行
        wcc_Head = wcc_Head & Format(CisFun.Val2(.発注区分), "0")      ' 発注区分
'Call Test_LogOutput(wcc_Head, 0)
'
    ''   * ボディ部作成 *
        wcc_Body = "": wcc_Body2 = ""
        For wNum1 = 0 To wRowSu                 ' 2008.09.12 upd : 10→wRowSu(15行)に変更
    '       * 品番 *
            If RTrim(.品番(wNum1)) = "" Then
                wcc_Body = wcc_Body & Space(18)
            Else
                wcc_Body = wcc_Body & Format(Mid(.品番(wNum1), 1, 18), "!" & String(18, "@"))
            End If

    '       * 背番号*
            If RTrim(.背番号(wNum1)) = "" Then
                wcc_Body = wcc_Body & Space(4)
            Else
                wcc_Body = wcc_Body & Format(Mid(.背番号(wNum1), 1, 4), "!@@@@")
            End If
    '       * 収容数 *
            wcc_Body = wcc_Body & Format(CisFun.Val2(.収容数(wNum1)), "00000")

    '       * 箱数 *
            wcc_Body = wcc_Body & Format(CisFun.Val2(.箱数(wNum1)), "000")

    '       * 納入数 *
            wcc_Body = wcc_Body & Format(CisFun.Val2(.納入数(wNum1)), "0000000")

        Next wNum1
    End With
'Call Test_LogOutput(wcc_Body, 0)
'
    If ChekData(0) Then
       H1lb_OCRCnt = Format(CisFun.Val2(H1lb_OCRCnt) + 1, "#,###")
       DoEvents
       QRProc = True
    Else
       Call QT10QSR_Close
       ProcHB = "B1"
       Call DispChange(ProcHB)
       Call SyoriName(SyoriKb)
    End If

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
    gSL_Select = gSL_Select & " ( 出荷指示書NO = 0"
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

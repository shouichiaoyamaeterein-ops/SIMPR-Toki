VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CSJ0020 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "材料使用報告書入力"
   ClientHeight    =   9960
   ClientLeft      =   3600
   ClientTop       =   1485
   ClientWidth     =   14730
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
   ScaleHeight     =   9960
   ScaleWidth      =   14730
   WindowState     =   2  '最大化
   Begin VB.PictureBox B1_Area1 
      Height          =   2115
      Left            =   450
      ScaleHeight     =   2055
      ScaleWidth      =   13800
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3135
      Width           =   13860
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   450
         Left            =   30
         Top             =   1590
         Width           =   13740
         _ExtentX        =   24236
         _ExtentY        =   794
         ForeColor       =   16711680
         Caption         =   "備　考"
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
         cPositionX      =   60
         Begin CisText_V60.CisText iB1_Biko 
            Height          =   345
            Left            =   945
            TabIndex        =   21
            Top             =   60
            Width           =   12930
            _ExtentX        =   22807
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   100
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   780
         Left            =   12180
         Top             =   810
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "在庫数"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Zaiko 
            Height          =   360
            Left            =   120
            Top             =   360
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "###,###"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cFont3D         =   2
            cAlingnment     =   4
            cPositionX      =   -60
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D23 
         Height          =   780
         Left            =   30
         Top             =   30
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   1376
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iB1_Hinbn 
            Height          =   360
            Left            =   75
            TabIndex        =   4
            Top             =   360
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D24 
         Height          =   780
         Left            =   3300
         Top             =   30
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "         材    質                 寸　　法"
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
         cPositionY      =   30
         Begin CisText_V60.CisText iB1_Zaist 
            Height          =   360
            Left            =   75
            TabIndex        =   5
            Top             =   360
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
         End
         Begin CisText_V60.CisText iB1_Itatu 
            Height          =   360
            Left            =   3180
            TabIndex        =   6
            Top             =   360
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
            cAlignment      =   1
         End
         Begin CisText_V60.CisText iB1_Width 
            Height          =   360
            Left            =   4125
            TabIndex        =   7
            Top             =   360
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
            cAlignment      =   1
         End
         Begin CisText_V60.CisText iB1_Long 
            Height          =   360
            Left            =   5070
            TabIndex        =   8
            Top             =   360
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
            cAlignment      =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D25 
         Height          =   780
         Index           =   0
         Left            =   10485
         Top             =   810
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "合計重量"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iB1_Gokei 
            Height          =   360
            Left            =   15
            TabIndex        =   20
            Top             =   360
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   635
            cFormat         =   "#########0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "#########0.00"
            MaxLength       =   13
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "#########0.00"
            cILength        =   10
            cAlignment      =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D25 
         Height          =   780
         Index           =   1
         Left            =   9405
         Top             =   30
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "加工数"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iB1_Kakou 
            Height          =   360
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   960
            _ExtentX        =   1693
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
            Text            =   "#######"
            MaxLength       =   7
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "#######"
            cILength        =   7
            cAlignment      =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D25 
         Height          =   780
         Index           =   2
         Left            =   10590
         Top             =   30
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "単重"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_TanJu 
            Height          =   360
            Left            =   105
            Top             =   360
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "###0.000"
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
            cFont3D         =   2
            cAlingnment     =   4
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D25 
         Height          =   780
         Index           =   3
         Left            =   11865
         Top             =   30
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "使用量"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iB1_Shiyo 
            Height          =   360
            Left            =   105
            TabIndex        =   10
            Top             =   345
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   635
            cFormat         =   "#########0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "#########0.00"
            MaxLength       =   13
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "#########0.00"
            cILength        =   10
            cAlignment      =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D25 
         Height          =   780
         Index           =   4
         Left            =   30
         Top             =   810
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "ラ ベ ル 重 量"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   345
            Left            =   45
            Top             =   360
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   609
            BackColor       =   16711935
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
            Begin CisText_V60.CisText iB1_LblH1 
               Height          =   285
               Left            =   2070
               TabIndex        =   12
               Top             =   30
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               cFormat         =   "##"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
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
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D8 
               Height          =   285
               Left            =   1890
               Top             =   30
               Width           =   180
               _ExtentX        =   318
               _ExtentY        =   503
               Caption         =   "×"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
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
            Begin CisText_V60.CisText iB1_LblJ1 
               Height          =   285
               Left            =   240
               TabIndex        =   11
               Top             =   30
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   503
               cFormat         =   "########0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "########0.000"
               MaxLength       =   13
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "########0.000"
               cILength        =   9
               cAlignment      =   1
            End
            Begin Cis3D_v60.CIS3D CIS3D19 
               Height          =   285
               Left            =   30
               Top             =   30
               Width           =   225
               _ExtentX        =   397
               _ExtentY        =   503
               BackColor       =   14737632
               Caption         =   "①"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   7
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D12 
            Height          =   345
            Left            =   2490
            Top             =   360
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            BackColor       =   16711935
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
            Begin CisText_V60.CisText iB1_LblJ2 
               Height          =   285
               Left            =   240
               TabIndex        =   13
               Top             =   30
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   503
               cFormat         =   "########0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "########0.000"
               MaxLength       =   13
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "########0.000"
               cILength        =   9
               cAlignment      =   1
            End
            Begin CisText_V60.CisText iB1_LblH2 
               Height          =   285
               Left            =   2055
               TabIndex        =   14
               Top             =   30
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               cFormat         =   "##"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
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
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D13 
               Height          =   285
               Left            =   1875
               Top             =   30
               Width           =   210
               _ExtentX        =   370
               _ExtentY        =   503
               Caption         =   "×"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   285
               Left            =   30
               Top             =   30
               Width           =   225
               _ExtentX        =   397
               _ExtentY        =   503
               BackColor       =   14737632
               Caption         =   "②"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   7
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D14 
            Height          =   345
            Left            =   4920
            Top             =   360
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            BackColor       =   16711935
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
            Begin CisText_V60.CisText iB1_LblH3 
               Height          =   285
               Left            =   2055
               TabIndex        =   16
               Top             =   30
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               cFormat         =   "##"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
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
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D16 
               Height          =   285
               Left            =   1875
               Top             =   30
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   503
               Caption         =   "×"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
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
            Begin CisText_V60.CisText iB1_LblJ3 
               Height          =   285
               Left            =   240
               TabIndex        =   15
               Top             =   30
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   503
               cFormat         =   "########0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "########0.000"
               MaxLength       =   13
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "########0.000"
               cILength        =   9
               cAlignment      =   1
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   285
               Left            =   30
               Top             =   30
               Width           =   225
               _ExtentX        =   397
               _ExtentY        =   503
               BackColor       =   14737632
               Caption         =   "③"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   7
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D26 
            Height          =   345
            Left            =   7350
            Top             =   360
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            BackColor       =   16711935
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
            Begin CisText_V60.CisText iB1_LblH4 
               Height          =   285
               Left            =   2055
               TabIndex        =   18
               Top             =   30
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               cFormat         =   "##"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
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
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D27 
               Height          =   285
               Left            =   1875
               Top             =   30
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   503
               Caption         =   "×"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
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
            Begin CisText_V60.CisText iB1_LblJ4 
               Height          =   285
               Left            =   240
               TabIndex        =   17
               Top             =   30
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   503
               cFormat         =   "########0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "########0.000"
               MaxLength       =   13
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "########0.000"
               cILength        =   9
               cAlignment      =   1
            End
            Begin Cis3D_v60.CIS3D CIS3D28 
               Height          =   285
               Left            =   30
               Top             =   30
               Width           =   225
               _ExtentX        =   397
               _ExtentY        =   503
               BackColor       =   14737632
               Caption         =   "④"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   7
               cBoderStyle     =   2
            End
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D25 
         Height          =   780
         Index           =   5
         Left            =   9840
         Top             =   810
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "戻し"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin VB.CheckBox Chk_Modosi 
            BackColor       =   &H00C0C0C0&
            Caption         =   "無"
            Height          =   270
            Left            =   75
            TabIndex        =   19
            Top             =   390
            Width           =   540
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   345
      Left            =   10260
      ScaleHeight     =   345
      ScaleWidth      =   3675
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   30
      Width           =   3675
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   345
         Index           =   1
         Left            =   2445
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BackColor       =   14737632
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
            Top             =   45
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   450
            BackColor       =   8421631
            ForeColor       =   0
            Caption         =   "ZZ9"
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
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   345
         Left            =   15
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BackColor       =   14737632
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
      Begin Cis3D_v60.CIS3D UC_3D7 
         Height          =   345
         Left            =   1230
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BackColor       =   14737632
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
            Top             =   45
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   450
            BackColor       =   16777088
            ForeColor       =   0
            Caption         =   "ZZ9"
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
      Height          =   1410
      Left            =   10710
      ScaleHeight     =   1350
      ScaleWidth      =   3510
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1065
      Width           =   3570
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   900
         Top             =   675
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
         Left            =   1755
         Top             =   675
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
         Left            =   2610
         Top             =   675
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
         Top             =   60
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   900
         Top             =   60
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
         BtnType         =   7
         BtnCaption      =   "新規"
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
         Left            =   1755
         Top             =   60
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   2610
         Top             =   60
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
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1320
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
   Begin VB.PictureBox B1_Area2 
      BackColor       =   &H00FF0000&
      Height          =   4500
      Left            =   450
      ScaleHeight     =   4440
      ScaleWidth      =   13800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5370
      Width           =   13860
      Begin Cis3D_v60.CIS3D BLB_DD 
         Height          =   4335
         Index           =   1
         Left            =   45
         Top             =   45
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   7646
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
         Begin vsFlexLib.vsFlexArray vsGrid 
            Height          =   4290
            Left            =   0
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   0
            Width           =   13635
            _Version        =   196608
            _ExtentX        =   24051
            _ExtentY        =   7567
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
            FormatString    =   $"CSJ0020.frx":0000
            Rows            =   11
            Cols            =   36
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   360
            SelectionMode   =   1
         End
         Begin CisText_V60.CisText B_TabDummy 
            Height          =   360
            Left            =   30
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   105
            Width           =   240
            _ExtentX        =   423
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
            Text            =   "X"
            MaxLength       =   1
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
      Caption         =   "【 材料使用報告書入力  】"
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
         Left            =   7635
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
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   345
      Index           =   3
      Left            =   0
      Top             =   0
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   270
         Index           =   12
         Left            =   390
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   30
         Value           =   1  'ﾁｪｯｸ
         Width           =   225
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   270
         Index           =   11
         Left            =   1065
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   30
         Value           =   1  'ﾁｪｯｸ
         Width           =   225
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   270
         Index           =   10
         Left            =   60
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   300
         Value           =   1  'ﾁｪｯｸ
         Width           =   225
      End
   End
   Begin Cis3D_v60.CIS3D H1_Area1 
      Height          =   885
      Left            =   480
      Top             =   1605
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   1561
      BackColor       =   16711680
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   765
         Left            =   1995
         Top             =   60
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1349
         Caption         =   "部門"
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
         cAlingnment     =   6
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   90
            TabIndex        =   2
            Top             =   345
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   345
            Left            =   1095
            Top             =   360
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   609
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   765
         Left            =   5730
         Top             =   60
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "作 業 者"
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
         cAlingnment     =   6
         cPositionY      =   50
         Begin Cis3D_v60.CIS3D H1lb_STannm 
            Height          =   345
            Left            =   645
            Top             =   360
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   609
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXX"
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
         Begin CisText_V60.CisText iH1_STanto 
            Height          =   360
            Left            =   120
            TabIndex        =   3
            Top             =   345
            Width           =   525
            _ExtentX        =   926
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
            Text            =   "XXX"
            MaxLength       =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   765
         Left            =   45
         Top             =   60
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   1349
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   50
         Begin CisYMD_With_Btn.CisYMDwB iH1_SYMD 
            Height          =   330
            Left            =   120
            TabIndex        =   1
            Top             =   345
            Width           =   1740
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
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.Label L_TestDisp 
      BackColor       =   &H00808000&
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   12000
      TabIndex        =   31
      Top             =   2640
      Width           =   2310
   End
End
Attribute VB_Name = "CSJ0020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   材料使用報告書入力
'**       フォームID    :   PSJ0020
'**       処理概要      :
'**
'**       作  成  日    :   2009/12/14  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid       As New CisVsGrid3
    Dim ActObj          As Object
    Dim ErrObj          As Object
'   実績区分の背景色
    Dim mColor_On       As Long
    Dim mColor_Off      As Long

'    Dim wTorKB          As Integer      ' 区分         ' 2009.12.17 del
    Dim wRowNo1         As Long
    Dim wRowNo2         As Long
    Dim wColNo1         As Long

'-------------------------------------------------------' 2009.12.17 del start
'    Dim wInput          As Boolean      ' 入力ﾌﾗｸﾞ
'
'    Dim wI              As Integer
'    Dim wSijiNo         As Long         '生産指示番号
'    Dim wRecNo          As Long
'    Dim wHinbn          As String       ' 品番
'    Dim wKoute          As String       ' 次工程
'    Dim wJyuCd          As String       ' 受注先CD
'    Dim wUke            As String       ' 受入
'    Dim wEda            As Byte         '枝番
'    Dim wBumon          As String       '部門CD
'    Dim wTanCD          As String       '担当者CD
'    Dim wSyuyo          As Long         '収容数
'    Dim wStopCD(40)     As String       '停止理由CD
'    Dim wAllAdd         As Boolean      '完全新規追加モード
'    Dim SiyoSu          As Long         '
'    Dim FuryoSiyoSu     As Long         '
'    Dim wJNoubn         As String
'    Dim wJKanriNo       As Long
'    Dim wJKanriNo1      As Long
'    Dim wJKanriNo2      As Long
'    Dim wJKanriNoZ      As Long
'    Dim wLastKotei      As Boolean      '最終工程判断フラグ
'    Dim mMaxKJno        As Long
'    Dim mMaxUpEda       As Integer
'
'    Dim wZSyuyo         As Currency
'-------------------------------------------------------' 2009.12.17 del end
    
    Dim mSL_Select      As String
    Dim mZSKanriNo      As Long             ' 新規材料使用量管理NO
    Dim mZRKanriNo      As Long             ' 材料管理番号
    Dim mZaikoSu        As Double           ' 在庫数
    Dim mSelectGyo      As Integer          ' 選択中の行
    Const mDebugFlg     As Boolean = False  ' デバッグ時はTrue
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub form_load()
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
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止

'    Call MeisyoInit         '名称マスタ初期設定            ' 2009.12.16 del
'    B1lb_Hinbn.Move iB1_Hinbn.Left, iB1_Hinbn.Top          ' 2009.12.14 del

    Call HeadBodyClear("H")
    Call HeadBodyClear("B")
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid
        .NoSpaceCol = 1
        .NumberSet = No
        .FixedHeight = 360
        .PatanMax = 1
'        .SelectCol = 32         ' 選択行マーク
        .SelectCol = 33         ' 選択行マーク
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = ""
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    iH1_SYMD = Format(Now(), "yyyyMMdd")
    iT1_Addsu = "": iT1_Mntsu = "": iT1_Delsu = ""
    L_TestDisp.Visible = mDebugFlg                  ' テスト表示領域 = デバッグモード時のみ有効
    
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
   'Static W_IVENT As String
   'W_IVENT = "Down"
    Key_Acc = False

    Select Case Key_Code
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
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

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
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
    Call ReturnPress
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
   Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   Call CisFrm.MousePT(11)
   
   mSelectGyo = 0
   
   If ProcHB = "H1" Then
      Call HeadBodyClear("H")
      iH1_SYMD = Format(Now(), "yyyyMMdd")
      GoTo PB_CAN_Ed
   End If
   
   If ProcHB = "B1" Then
      Call HeadBodyClear("B")
      Call GridNoSelect
      
      If SyoriKB = "MNT" Then
         ProcHB = "B2"
         Call DispChange(ProcHB)
         Call SyoriName(SyoriKB)
         PB_ENT.Visible = False
         Call VSGrid_Click
         Call vsGrid.SetFocus
      Else
         Call GridClear
         ProcHB = "H1"
         Call DispChange(ProcHB)
         SyoriKB = ""
         Call SyoriName(SyoriKB)
      End If
      GoTo PB_CAN_Ed
   End If

   If ProcHB = "B2" Then
      Call HeadBodyClear("B")
      Call GridNoSelect
      Call GridClear
      ProcHB = "H1"
      Call DispChange(ProcHB)
      SyoriKB = ""
      Call SyoriName(SyoriKB)
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
    Set ActObj = Dummy
    
    If SyoriKB <> "ADD" Then
        Dummy.Enabled = True: Dummy.SetFocus: DoEvents
        Call GridNoSelect
        ProcHB = "B1"
        Call DispChange(ProcHB)
        SyoriKB = "ADD"
        Call HeadBodyClear("B")
      ''Call HeadBodyClear("H")
        Call SyoriName(SyoriKB)
        iB1_Hinbn.SetFocus
        B1_Area2.Enabled = True: B1_Area2.BackColor = gPGotSel
        vsGrid.HighLight = flexHighlightNever   ' 新規時はグリッドの現在行色をなくす
        mSelectGyo = 0
        Dummy.Enabled = False
    End If
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    Dim wCnt    As Integer
    Set ActObj = vsGrid
    
    If Trim(vsGrid.TextMatrix(1, 1)) = "" Then
        With CisFun
            .MB_Title = ""
            .MB_Button = Error
            .MB_Lines = 4
            .MB_MSG(2) = "    修正対象データはありません            "
            .MBOX
            Exit Sub
        End With
    End If
    
    If SyoriKB <> "MNT" Then
        Dummy.Enabled = True: Dummy.SetFocus: DoEvents
        Call HeadBodyClear("B")
        Call GridNoSelect
        If SyoriKB = "ADD" Then Call HeadBodyClear("B")
        ProcHB = "B2"
        For wCnt = 1 To vsGrid.Rows - 1
            If Trim(vsGrid.TextMatrix(wCnt, CisVsGrid.SelectCol)) <> "" Then
                ProcHB = "B1"
                Exit For
            End If
        Next
        Call DispChange(ProcHB)
        B1_Area2.Enabled = True
        B1_Area2.BackColor = gPGotSel
        If ProcHB = "B1" Then iB1_Hinbn.SetFocus
''        If ProcHB = "B2" Then PB_ENT.Visible = False
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
        Dummy.Enabled = False
        
        vsGrid.HighLight = flexHighlightWithFocus   ' 新規以外はグリッドの現在行色を表示する
        mSelectGyo = 0
        vsGrid.TopRow = 1
        vsGrid.Row = 1
        Call VSGrid_Click
        Call vsGrid.SetFocus
    End If
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    Dim wCnt    As Integer
    Set ActObj = vsGrid
    
    If Trim(vsGrid.TextMatrix(1, 1)) = "" Then
        With CisFun
            .MB_Title = ""
            .MB_Button = Error
            .MB_Lines = 4
            .MB_MSG(2) = "    削除対象データはありません            "
            .MBOX
            Exit Sub
        End With
    End If
    
    If SyoriKB <> "DEL" Then
        Dummy.Enabled = True: Dummy.SetFocus: DoEvents
        Call HeadBodyClear("B")
        Call GridNoSelect
        ProcHB = "B2"
        Call DispChange(ProcHB)
        SyoriKB = "DEL"
        Call SyoriName(SyoriKB)
        Dummy.Enabled = False
        
        vsGrid.HighLight = flexHighlightWithFocus   ' 新規以外はグリッドの現在行色を表示する
        mSelectGyo = 0
        vsGrid.TopRow = 1
        vsGrid.Row = 1
        Call VSGrid_Click
        Call vsGrid.SetFocus
        '''PB_ENT.Visible = False                      ' 選択されるまで入力ボタンは非表示
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iB1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iB1_Zaist" Then Call Look_Zaist: Exit Sub
    If PB_Look.Tag = "iB1_Itatu" Then Call Look_Zaist: Exit Sub
    If PB_Look.Tag = "iB1_Width" Then Call Look_Zaist: Exit Sub
    If PB_Look.Tag = "iB1_Long" Then Call Look_Zaist: Exit Sub
    If PB_Look.Tag = "iH1_STanto" Then Call Look_Tanto: Exit Sub

    If PB_Look.Tag = "iH1_SYMD" Then
       iH1_SYMD.ShowCalender
       Exit Sub
    End If
End Sub
'+----------------------------+
'+       名称検索 - 担当      +
'+----------------------------+
Private Sub Look_Tanto()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RVI_Kubun = "担当管理"

    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
       iH1_STanto = RTrim(RV_Code)
       H1lb_STannm = RTrim(RV_Kubun)
    End If
End Sub
'+----------------------------+
'+       部門検索             +
'+----------------------------+
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
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
   'RV_Call = ""
    RVI_Hinmoku = 9
    RVI_SK = 9

    RVI_Torcd = Trim(iH1_Torcd)
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iB1_Hinbn = RTrim(RV_Hinbn)
    End If
End Sub
''+------------------------+
''+     材料取引先検索     +
''+------------------------+
Private Sub Look_Zaist()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 1
    RVI_SK = 9
    RVI_Tehai = ""
    RVI_Torcd = ""

    CKK0077.Show vbModal
    Unload CKK0077
    Set CKK0077 = Nothing

    If RV_Rtn Then
       iB1_Zaist = RV_Zaist
       iB1_Itatu = RV_Itatu
       iB1_Width = RV_Width
       iB1_Long = RV_Long
    End If
End Sub
'+-----------------------------+
'+
'+-----------------------------+
Private Sub PNL_PF_Click()
   'On Error Resume Next
    ActObj.SetFocus
   'On Error GoTo 0
End Sub

'+-----------------------------+
'+  iH1_Torcd   (部門)
'+-----------------------------+
Private Sub iH1_Torcd_GotFocus()
    Set ActObj = iH1_Torcd
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    H1lb_Tornm = ""
    If Trim(iH1_Torcd) <> "" Then
        gSL_Select = "SELECT * FROM 取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先CD = '" & Trim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 社内区分 = 1"
        If TRMRead(gSL_Select) Then
            H1lb_Tornm = TRM.正式名称
        Else
            H1lb_Tornm = ""
        End If
    End If
End Sub
'+-----------------------------+
'+  iH1_SYMD
'+-----------------------------+
Private Sub iH1_SYMD_GotFocus()
    Set ActObj = iH1_SYMD
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_SYMD_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+      iH1_STanto
'+-----------------------------+
Private Sub iH1_STanto_GotFocus()
    Set ActObj = iH1_STanto
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_STanto_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_STannm = ""
    If RTrim(iH1_STanto) <> "" Then
       If MeisyoGet("担当管理", iH1_STanto) Then
          H1lb_STannm = SYM_Meisyo
       End If
    End If
End Sub
'+-----------------------------+
'+      iB1_Hinbn
'+-----------------------------+
Private Sub iB1_Hinbn_GotFocus()
    Set ActObj = iB1_Hinbn
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    If Trim(iB1_Zaist) = "" And Trim(iB1_Itatu) = "" And Trim(iB1_Width) = "" And Trim(iB1_Long) = "" Then
        Call GetZaist
    End If
    Call GetTanju
    Call DispZaikoSu
End Sub
'+-----------------------------+
'+      iB1_Zaist
'+-----------------------------+
Private Sub iB1_Zaist_GotFocus()
    Set ActObj = iB1_Zaist
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Zaist_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    If Trim(iB1_Itatu) = "" And Trim(iB1_Width) = "" And Trim(iB1_Long) = "" Then
        Call GetZaist
    End If
    If RTrim(B1lb_TanJu) = "" And RTrim(iB1_Zaist) <> "" Then
       Call GetZairyo
    End If
    Call GetTanju
    Call DispZaikoSu            ' Call GetZaikoSu
End Sub
'+-----------------------------+
'+      iB1_Itatu
'+-----------------------------+
Private Sub iB1_Itatu_GotFocus()
    Set ActObj = iB1_Itatu
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Itatu_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    If RTrim(B1lb_TanJu) = "" And RTrim(iB1_Itatu) <> "" Then
       Call GetZairyo
    End If
    Call GetTanju
    Call DispZaikoSu            ' Call GetZaikoSu
End Sub
'+-----------------------------+
'+      iB1_Width
'+-----------------------------+
Private Sub iB1_Width_GotFocus()
    Set ActObj = iB1_Width
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Width_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    If RTrim(B1lb_TanJu) = "" And RTrim(iB1_Width) <> "" Then
       Call GetZairyo
    End If
    Call GetTanju
    Call DispZaikoSu            ' Call GetZaikoSu
End Sub
'+-----------------------------+
'+      iB1_Long
'+-----------------------------+
Private Sub iB1_Long_GotFocus()
    Set ActObj = iB1_Long
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Long_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    If RTrim(B1lb_TanJu) = "" And RTrim(iB1_Long) <> "" Then
       Call GetZairyo
    End If
    Call GetTanju
    Call DispZaikoSu            ' Call GetZaikoSu
End Sub
'+-----------------------------+
'+      単重抽出(＋材料管理番号取得）
'+-----------------------------+
Private Sub GetTanju()
    gSL_Select = ""
'    gSL_Select = gSL_Select & "SELECT ZM.* FROM  材料マスタ ZM"
'    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 材料品番マスタ ZH"
'    gSL_Select = gSL_Select & "    ON ZM.材料管理番号 = ZH.材料管理番号"
'    gSL_Select = gSL_Select & " WHERE ZH.使用品番 = '" & Trim(iB1_Hinbn) & "'"
'    gSL_Select = gSL_Select & "   AND ZM.材質     = '" & Trim(iB1_Zaist) & "'"
'    gSL_Select = gSL_Select & "   AND ZM.板厚     =  " & CisFun.Val2(iB1_Itatu) & " "
'    gSL_Select = gSL_Select & "   AND ZM.幅       =  " & CisFun.Val2(iB1_Width) & " "
'    gSL_Select = gSL_Select & "   AND ZM.長さ     =  " & CisFun.Val2(iB1_Long) & "  "
'    If ZRMRead(gSL_Select) Then
'       B1lb_TanJu = Format(ZRM.単位, "###0.000")
'       mZRKanriNo = ZHM.材料管理番号
'    Else
'       B1lb_TanJu = ""
'       mZRKanriNo = 0
'    End If
    gSL_Select = gSL_Select & "SELECT * FROM 材料品番マスタ"
    gSL_Select = gSL_Select & " WHERE REPLACE(使用品番,'-','') ='" & Trim(iB1_Hinbn) & "'"
    If ZHMRead(gSL_Select) Then
       B1lb_TanJu = Format(ZHM.使用量, "###0.000")
       mZRKanriNo = ZHM.材料管理番号
    Else
       B1lb_TanJu = ""
'       mZRKanriNo = 0
    End If
End Sub
''+-----------------------------+
''+      在庫数抽出
''+-----------------------------+
'Private Sub GetZaikoSu()
'    gSL_Select = ""
'    gSL_Select = gSL_Select & "SELECT ZZ.* FROM  材料在庫マスタ ZZ"
'    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 材料マスタ ZM"
'    gSL_Select = gSL_Select & "    ON ZM.材料管理番号 = ZZ.材料管理番号"
'    gSL_Select = gSL_Select & " WHERE ZM.材質   = '" & Trim(iB1_Zaist) & "'"
'    gSL_Select = gSL_Select & "   AND ZM.板厚   =  " & CisFun.Val2(iB1_Itatu) & " "
'    gSL_Select = gSL_Select & "   AND ZM.幅     =  " & CisFun.Val2(iB1_Width) & " "
'    gSL_Select = gSL_Select & "   AND ZM.長さ   =  " & CisFun.Val2(iB1_Long) & "  "
'    gSL_Select = gSL_Select & "   AND ZZ.年月度 = '" & Mid(iH1_SYMD, 1, 6) & "'"
'    B1lb_Zaiko = ""
'    If ZZMRead(gSL_Select) Then
'       If ZZM.在庫数 <> 0 Then
'          B1lb_Zaiko = Format(ZZM.在庫数, "###0.000")
'       End If
'    End If
'End Sub
'+-----------------------------+
'+      材料抽出
'+-----------------------------+
Private Sub GetZaist()
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT ZM.* FROM  材料マスタ ZM"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 材料品番マスタ ZH"
    gSL_Select = gSL_Select & "    ON ZH.材料管理番号 = ZM.材料管理番号"
    gSL_Select = gSL_Select & " WHERE ZH.使用品番 = '" & Trim(iB1_Hinbn) & "'"
    If ZRMRead(gSL_Select) Then
        iB1_Zaist = ZRM.材質
        iB1_Itatu = ZRM.板厚
        iB1_Width = ZRM.幅
        iB1_Long = ZRM.長さ
    End If
End Sub
'+-----------------------------+
'+      材料マスタ検索
'+-----------------------------+
Private Sub GetZairyo()
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT 材料管理番号 FROM  材料マスタ "
    gSL_Select = gSL_Select & " WHERE 材質 = '" & Trim(iB1_Zaist) & "'"
    gSL_Select = gSL_Select & " AND   板厚 = " & CisFun.Val2(iB1_Itatu)
    gSL_Select = gSL_Select & " AND   幅 = " & CisFun.Val2(iB1_Width)
    gSL_Select = gSL_Select & " AND   長さ = " & CisFun.Val2(iB1_Long)
    If ZRMRead(gSL_Select) Then
        mZRKanriNo = ZRM.材料管理番号
    End If
End Sub
'+-----------------------------+
'+      iB1_Kakou
'+-----------------------------+
Private Sub iB1_Kakou_GotFocus()
    Set ActObj = iB1_Kakou
End Sub
Private Sub iB1_Kakou_LostFocus()
    Dim wKaz As Double
    
    If Trim(iB1_Shiyo) = "" Then
        wKaz = CisFun.Val2(iB1_Kakou) * CisFun.Val2(B1lb_TanJu)
        If wKaz <> 0 Then
            iB1_Shiyo = Format(wKaz, "###0.00")
        End If
    End If
End Sub
'+-----------------------------+
'+      B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
'-------------------------------------------------------' 2009.12.14 del start
'    If B_TabDummy.Tag = "S" Then
'        Call TextToVsGrid(Val(Input_Area.Tag))
'        Call Grid_RowChange(vbKeyUp)
'    End If
''    If B_TabDummy.Tag = "E" Then
''        Call TextToVsGrid(Val(Input_Area.Tag))
''        Call Grid_RowChange(vbKeyDown)
''    End If
'    SendKeys "{Tab}"
'-------------------------------------------------------' 2009.12.14 del end
End Sub
'+-----------------------------+
'+      iB1_Modosi (戻し)
'+-----------------------------+
Private Sub Chk_Modosi_Click()
    If Chk_Modosi.Value = 0 Then
       Chk_Modosi.Caption = "無"
    Else
       Chk_Modosi.Caption = "有"
    End If
End Sub
Private Sub Chk_Modosi_GotFocus()
    Set ActObj = Chk_Modosi
End Sub
'+-----------------------------+
'+      フォーカス記憶のみ
'+-----------------------------+
Private Sub iB1_Shiyo_GotFocus()
    Set ActObj = iB1_Shiyo
End Sub
Private Sub iB1_LblJ1_GotFocus()
    Set ActObj = iB1_LblJ1
End Sub
Private Sub iB1_LblH1_GotFocus()
    Set ActObj = iB1_LblH1
End Sub
Private Sub iB1_LblJ2_GotFocus()
    Set ActObj = iB1_LblJ2
End Sub
Private Sub iB1_LblH2_GotFocus()
    Set ActObj = iB1_LblH2
End Sub
Private Sub iB1_LblJ3_GotFocus()
    Set ActObj = iB1_LblJ3
End Sub
Private Sub iB1_LblH3_GotFocus()
    Set ActObj = iB1_LblH3
End Sub
Private Sub iB1_LblJ4_GotFocus()
    Set ActObj = iB1_LblJ4
End Sub
Private Sub iB1_LblH4_GotFocus()
    Set ActObj = iB1_LblH4
End Sub
Private Sub iB1_Gokei_GotFocus()
    Set ActObj = iB1_Gokei
End Sub

'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Dim W_Color     As Long
    W_Color = &H80000008 '初期色ｾｯﾄ(黒)
    Select Case SYKB
        Case "ADD"
            SyoriNM = "追加"
            K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
        Case "MNT"
            SyoriNM = "修正"
            K_Sykbnm.ForeColor = &HFF0000  '(青色)
        Case "DEL"
            SyoriNM = "削除"
            K_Sykbnm.ForeColor = &HFF&     '(赤色)
            W_Color = &HFF&
        Case Else
            SyoriNM = ""
    End Select
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
        If gObj.Name Like "*vsGrid*" Then gObj.ForeColor = W_Color
        If ProcHB <> "H1" Then
    '       * PF-Key 表示設定 *
                If gObj.Name = "PB_ADD" Or _
                   gObj.Name = "PB_MNT" Or _
                   gObj.Name = "PB_DEL" Then
                   gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SYKB)
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
    Call CisFrm.TimeSet(1)
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
    H1_Area1.Enabled = H1Mode:  H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode:  B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B2Mode:  B1_Area2.BackColor = B2Color
    
    PB_Look.Visible = False
    PB_ADD.Visible = False
    PB_MNT.Visible = False
    PB_DEL.Visible = False
    PB_CAN.Visible = True
    PB_ENT.Visible = True
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
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = "": .Enabled = True
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False: .BackColor = mColor_Off
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Cmb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        Chk_Modosi.Value = 0
        mZRKanriNo = 0
    End If
End Sub

'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        Call PB_ADD_Click
        Call DBInput("")
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B1" Then
'       << 新規・修正処理 >>
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call GridNoSelect
        Call GridClear
        Call HeadBodyClear("B")
        
        gSL_Select = mSL_Select     ' 退避していたSQL セット
        Call DBInput("")            ' 再読み込み
        If SyoriKB = "ADD" Then
           SyoriKB = "": Call SyoriName("")
           Call PB_ADD_Click
           vsGrid.TopRow = vsGrid.Rows - 10
        Else
           SyoriKB = "": Call SyoriName("")
           Call PB_MNT_Click
        End If
        GoTo ReturnPress_Ed
    End If

    If ProcHB = "B2" Then
        Call VSGrid_DblClick
'       << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not SelectChk Then
               Call vsGrid.SetFocus
               GoTo ReturnPress_Ed
            End If
            
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call GridNoSelect
            Call GridClear
            Call HeadBodyClear("B")
            
            gSL_Select = mSL_Select     ' 退避していたSQL セット
            If Not DBInput("") Then     ' 再表示時にデータが無くなっていたら新規に変更する
                ProcHB = "B1"
                Call DispChange(ProcHB)
                SyoriKB = "ADD"
                Call SyoriName(SyoriKB)
                B1_Area2.Enabled = True: B1_Area2.BackColor = gPGotSel
                GoTo ReturnPress_Ed
            End If
            SyoriKB = "": Call SyoriName("")
            Call PB_DEL_Click
        Else
            Call VSGrid_DblClick
        End If
        GoTo ReturnPress_Ed
    End If

ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      行選択ﾁｪｯｸ処理
'****************************
Private Function SelectChk() As Boolean
    Dim wCnt    As Integer
    SelectChk = False
    
    With vsGrid
        For wCnt = 1 To .Rows - 1
            If Trim(.TextMatrix(wCnt, CisVsGrid.SelectCol)) <> "" Then Exit For
        Next
        If wCnt >= .Rows Then
            With CisFun
                .MB_Lines = 4
                .MB_Title = ""
                .MB_Button = Error
                .MB_MSG(2) = "    行を選択してください           "
                .MBOX
                Exit Function
            End With
        End If
    End With
    
    SelectChk = True
End Function
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

' 作業日
    If Trim(iH1_SYMD) = "" Or Len(Trim(iH1_SYMD)) < 8 Then
        CisFun.MB_MSG(2) = "    作業日を入力して下さい            "
        CisFun.MBOX
        iH1_SYMD.SetFocus
        Exit Function
    End If
' 部門
    If Trim(iH1_Torcd) = "" Then
        CisFun.MB_MSG(2) = "    部門を入力して下さい            "
        CisFun.MBOX
        iH1_Torcd.SetFocus
        Exit Function
    End If

' 担当者
    If Trim(iH1_STanto) <> "" Then
       If MeisyoGet("担当管理", iH1_STanto) Then
          H1lb_STannm = SYM_Meisyo
       Else
          CisFun.MB_MSG(2) = "    担当者コード未登録        "
          CisFun.MBOX
          iH1_STanto.SetFocus
          Exit Function
       End If
    End If


' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT ZS.* "
    gSL_Select = gSL_Select & " FROM  材料使用量テーブル ZS"
    gSL_Select = gSL_Select & " WHERE ZS.作業日 = '" & iH1_SYMD & "'"
    gSL_Select = gSL_Select & "   AND ZS.部門   = '" & Trim(iH1_Torcd) & "'"
    If Trim(iH1_STanto) <> "" Then
       gSL_Select = gSL_Select & "   AND ZS.作業者 = '" & Trim(iH1_STanto) & "'"
    End If
    gSL_Select = gSL_Select & " ORDER BY ZS.材料使用量管理NO, ZS.材料使用量管理行"

' 次回同データ読み込み時用に保存
    mSL_Select = gSL_Select
    
Head1Chk_ED:
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
'
    Dim wKanriNo As Long

'ﾒｯｾｰｼﾞﾎﾞｯｸｽ設定
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
    
    If B1_Area1.Enabled = False Then
       If SyoriKB = "MNT" Then
          CisFun.MB_MSG(2) = "    修正する行を選択して下さい            "
       Else
          CisFun.MB_MSG(2) = "    修正する行を選択して下さい            "
       End If
       CisFun.MBOX
       Exit Function
    End If
    
    If Trim(iB1_Hinbn) = "" Then
        CisFun.MB_MSG(2) = "    品番を入力して下さい            "
        CisFun.MBOX
        iB1_Hinbn.SetFocus
        Exit Function
    End If
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT * FROM 品番マスタ WHERE 品番 = '" & Trim(iB1_Hinbn) & "'"
    If Not HNMRead(gSL_Select) Then
        CisFun.MB_MSG(2) = "    品番　未登録            "
        CisFun.MBOX
        iB1_Hinbn.SetFocus
        Exit Function
    End If
    
    If Trim(iB1_Zaist) = "" Then
        CisFun.MB_MSG(2) = "    材質を入力して下さい            "
        CisFun.MBOX
        iB1_Zaist.SetFocus
        Exit Function
    End If
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT * FROM 材料マスタ"
    gSL_Select = gSL_Select & " WHERE 材質 = '" & Trim(iB1_Zaist) & "'"
    gSL_Select = gSL_Select & "   AND 板厚 =  " & CisFun.Val2(iB1_Itatu)
    gSL_Select = gSL_Select & "   AND 幅   =  " & CisFun.Val2(iB1_Width)
    gSL_Select = gSL_Select & "   AND 長さ =  " & CisFun.Val2(iB1_Long)
    If Not ZRMRead(gSL_Select) Then
        CisFun.MB_MSG(2) = "    材料　未登録            "
        CisFun.MBOX
        iB1_Zaist.SetFocus
        Exit Function
    End If
    wKanriNo = ZRM.材料管理番号
    
    mZRKanriNo = ZRM.材料管理番号
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT * FROM 材料品番マスタ"
    gSL_Select = gSL_Select & " WHERE 材料管理番号 = " & Str(wKanriNo) & ""
    gSL_Select = gSL_Select & "   AND 使用品番     = '" & Trim(iB1_Hinbn) & "'"
    If Not ZHMRead(gSL_Select) Then
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(2) = "    材料品番マスタが未登録です                    "
        CisFun.MB_MSG(4) = "    よろしいですか？　　　　　　　　　            "
        CisFun.MB_Button = OK_CAN
        If Not CisFun.MBOX Then
            iB1_Zaist.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(iB1_Kakou) = "" Then
        CisFun.MB_MSG(2) = "    加工数を入力して下さい            "
        CisFun.MBOX
        iB1_Kakou.SetFocus
        Exit Function
    End If
    
    If Trim(iB1_Shiyo) = "" Then
        iB1_Shiyo = Format(CisFun.Val2(iB1_Kakou) * CisFun.Val2(B1lb_TanJu), "###0.00")
    Else
        If CisFun.Val2(iB1_Shiyo) <> (CisFun.Val2(iB1_Kakou) * CisFun.Val2(B1lb_TanJu)) Then
            CisFun.MB_Lines = 5
            CisFun.MB_MSG(2) = "    使用量が加工数×単重と一致しません            "
            CisFun.MB_MSG(4) = "    よろしいですか？　　　　　　　　　            "
            CisFun.MB_Button = OK_CAN
            If Not CisFun.MBOX Then
                iB1_Shiyo.SetFocus
                Exit Function
            End If
        End If
    End If
    
    Call DispZaikoSu
    If CisFun.Val2(B1lb_Zaiko) - CisFun.Val2(iB1_Shiyo) < 0 Then
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(2) = "    現在在庫数－使用量がゼロ以下です              "
        CisFun.MB_MSG(4) = "    よろしいですか？　　　　　　　　　            "
        CisFun.MB_Button = OK_CAN
        If Not CisFun.MBOX Then
            iB1_Shiyo.SetFocus
            Exit Function
        End If
    End If
    
    Body1Chk = True
    Exit Function
Body1Chk_Err:
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

    If Not GridSet Then Exit Function
    DBInput = True
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
'+----------------------+
'+      削 除 処 理      +
'+----------------------+
Private Function DBDelete(Optional DBNo As Integer) As Boolean
    Dim wCnt    As Integer
    
    DBDelete = False

' ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "    削除処理を行います。            "
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Title = "削除確認"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With

    For wCnt = 1 To vsGrid.Rows - 1
        If Trim(vsGrid.TextMatrix(wCnt, CisVsGrid.SelectCol) <> "") Then Exit For
    Next
    If wCnt >= vsGrid.Rows Then GoTo DBDelete_Err
    
    On Error GoTo DBDelete_Err
    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 削除処理 )
    gSL_Select = ""
    gSL_Select = gSL_Select & "DELETE FROM 材料使用量テーブル"
    gSL_Select = gSL_Select & " WHERE 材料使用量管理NO = " & Trim(vsGrid.TextMatrix(wCnt, 26))
    gSL_Select = gSL_Select & "   AND 材料使用量管理行 = " & Trim(vsGrid.TextMatrix(wCnt, 27))
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then GoTo DBDelete_Err
    End With
    
'   件数表示
    iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")

    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    On Error GoTo 0

    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    Dim wKNo    As Integer
    Dim wGyo    As Integer
    Dim wCnt    As Integer
    
    DBPut = False

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True

' ( 確認メッセージ )
    With CisFun
        .MB_Lines = 5
        If SyoriKB = "ADD" Then
            .MB_MSG(2) = "    追加処理を行います。            "
            .MB_Title = "追加確認"
        Else
            .MB_MSG(2) = "    修正処理を行います。            "
            .MB_Title = "修正確認"
        End If
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Exit Function
        End If
    End With
        
    If SyoriKB = "ADD" Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT 材料使用量管理NO, max(材料使用量管理行) as 材料使用量管理行"
        gSL_Select = gSL_Select & " FROM  材料使用量テーブル"
        gSL_Select = gSL_Select & " WHERE 作業日 = '" & iH1_SYMD & "'"
        gSL_Select = gSL_Select & "   AND 部門   = '" & Trim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " GROUP BY 材料使用量管理NO"
        If ZSRRead(gSL_Select) Then
           wKNo = ZSR.材料使用量管理NO
           wGyo = ZSR.材料使用量管理行 + 1
        Else
           Call GetKanriNo      ' 新規材料使用量管理NO取得
           wKNo = mZSKanriNo
           wGyo = 1
        End If
    Else
        For wCnt = 1 To vsGrid.Rows - 1
            If Trim(vsGrid.TextMatrix(wCnt, CisVsGrid.SelectCol)) <> "" Then Exit For
        Next
        wKNo = vsGrid.TextMatrix(wCnt, 26)
        wGyo = vsGrid.TextMatrix(wCnt, 27)
    End If

    Call ItemsClearZSR
    With ZSR
        .材料使用量管理NO = wKNo
        .材料使用量管理行 = wGyo
        If SyoriKB = "ADD" Then
           .作業日 = iH1_SYMD
           .部門 = Trim(iH1_Torcd)
           .作業者 = Trim(iH1_STanto)
        Else
           .作業日 = Trim(vsGrid.TextMatrix(wCnt, 28))
           .部門 = Trim(vsGrid.TextMatrix(wCnt, 29))
           .作業者 = Trim(vsGrid.TextMatrix(wCnt, 30))
        End If
        .品番 = Trim(iB1_Hinbn)
        .材質 = Trim(iB1_Zaist)
        .板厚 = CisFun.Val2(iB1_Itatu)
        .幅 = CisFun.Val2(iB1_Width)
        .長さ = CisFun.Val2(iB1_Long)
        .材料管理番号 = mZRKanriNo
        .加工数 = CisFun.Val2(iB1_Kakou)
        .単重 = CisFun.Val2(B1lb_TanJu)
        .使用量 = CisFun.Val2(iB1_Shiyo)
        .ラベル重量1 = CisFun.Val2(iB1_LblJ1)
        .ラベル本数1 = CisFun.Val2(iB1_LblH1)
        .ラベル重量2 = CisFun.Val2(iB1_LblJ2)
        .ラベル本数2 = CisFun.Val2(iB1_LblH2)
        .ラベル重量3 = CisFun.Val2(iB1_LblJ3)
        .ラベル本数3 = CisFun.Val2(iB1_LblH3)
        .ラベル重量4 = CisFun.Val2(iB1_LblJ4)
        .ラベル本数4 = CisFun.Val2(iB1_LblH4)
        .戻し重量 = Chk_Modosi.Value
        .合計重量 = CisFun.Val2(iB1_Gokei)
        
        .備考 = RTrim(iB1_Biko)
        
    End With

    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    If SyoriKB = "ADD" Then
       If Not ZSRInsert Then GoTo DBPut_Err
    Else
       If Not ZSRUpdate Then GoTo DBPut_Err
    End If
    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

'   件数表示
    Select Case SyoriKB
        Case "ADD": iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
        Case "MNT": iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End Select

    On Error GoTo 0

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
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
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 4, 0, "No.")
        Call .FixedSet(中中, 左中, あり, 26, 1, "品　番")
        Call .FixedSet(中中, 左中, あり, 26, 2, "材　質")
        Call .FixedSet(中中, 左中, あり, 20, 3, "寸　法")
        Call .FixedSet(中中, 右中, あり, 8, 4, "加工数")
        Call .FixedSet(中中, 右中, あり, 8, 5, "単重")
        Call .FixedSet(中中, 右中, なし, 14, 6, "使用量")
        Call .FixedSet(中中, 左中, なし, 16, 7, "ラベル重量1")
        Call .FixedSet(中中, 左中, なし, 16, 8, "ラベル重量2")
        Call .FixedSet(中中, 左中, なし, 16, 9, "ラベル重量3")
        Call .FixedSet(中中, 左中, なし, 16, 10, "ラベル重量4")
        Call .FixedSet(中中, 中中, なし, 4, 11, "戻し")
        Call .FixedSet(中中, 右中, なし, 14, 12, "合計重量")
       'If mDebugFlg Then
           Call .FixedSet(中中, 左中, なし, 8, 13, "材管番号")
           Call .FixedSet(中中, 右中, なし, 8, 14, "板厚")
           Call .FixedSet(中中, 右中, なし, 8, 15, "幅")
           Call .FixedSet(中中, 右中, なし, 8, 16, "長さ")
           Call .FixedSet(中中, 右中, なし, 8, 17, "重量 1")
           Call .FixedSet(中中, 右中, なし, 8, 18, "本数 1")
           Call .FixedSet(中中, 右中, なし, 8, 19, "重量 2")
           Call .FixedSet(中中, 右中, なし, 8, 20, "本数 2")
           Call .FixedSet(中中, 右中, なし, 8, 21, "重量 3")
           Call .FixedSet(中中, 右中, なし, 8, 22, "本数 3")
           Call .FixedSet(中中, 右中, なし, 8, 23, "重量 4")
           Call .FixedSet(中中, 右中, なし, 8, 24, "本数 4")
           Call .FixedSet(中中, 右中, なし, 8, 25, "戻し重量")
           Call .FixedSet(中中, 右中, なし, 8, 26, "管理NO")
           Call .FixedSet(中中, 右中, なし, 6, 27, "管理行")
           Call .FixedSet(中中, 左中, なし, 8, 28, "作業日")
           Call .FixedSet(中中, 左中, なし, 7, 29, "部門")
           Call .FixedSet(中中, 左中, なし, 6, 30, "作業者")
           Call .FixedSet(中中, 右中, なし, 2, 31, "31")
           Call .FixedSet(中中, 中中, なし, 8, 32, "Select")
           Call .FixedSet(中中, 右中, なし, 2, 33, "33")
        Call .FixedSet(中中, 左中, なし, 100, 34, "備　　　考")
       'End If
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'        Call .FixedSet(中中, 左中, なし, 0, 10, "")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12")
        If mDebugFlg Then Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33")
        .InitDisp
    End With
    
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Function GridSet() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
    Dim wStr        As String

    GridSet = False

    If ZSRRead(gSL_Select) Then

'       抽出処理 初期ｾｯﾄ
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
'       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With vsGrid
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
                .Rows = CisVsGrid.DispMax
            Else
                .Rows = Ttl_Kensu + 1
            End If
        End With
        
        ID = 1
        Do Until Not ZSR_RDSTS
            With vsGrid
                .TextMatrix(ID, 0) = ID
                .TextMatrix(ID, 1) = ZSR.品番
                .TextMatrix(ID, 2) = ZSR.材質
                .TextMatrix(ID, 3) = Format(ZSR.板厚, "###0.00") & " x " & Format(ZSR.幅, "###0.00") & " x " & Format(ZSR.長さ, "###0.00")
                .TextMatrix(ID, 4) = Format(ZSR.加工数, "####")
                .TextMatrix(ID, 5) = Format(ZSR.単重, "###0.000")
                .TextMatrix(ID, 6) = Format(ZSR.使用量, "#########0.00")
                .TextMatrix(ID, 7) = Format(ZSR.ラベル重量1, "########0.000") & "x" & Format(ZSR.ラベル本数1, "##")
                .TextMatrix(ID, 8) = Format(ZSR.ラベル重量2, "########0.000") & "x" & Format(ZSR.ラベル本数2, "##")
                .TextMatrix(ID, 9) = Format(ZSR.ラベル重量3, "########0.000") & "x" & Format(ZSR.ラベル本数3, "##")
                .TextMatrix(ID, 10) = Format(ZSR.ラベル重量4, "########0.000") & "x" & Format(ZSR.ラベル本数4, "##")
                wStr = "無"
                If ZSR.戻し重量 <> 0 Then
                   wStr = "有"
                End If
                .TextMatrix(ID, 11) = Format(ZSR.戻し重量, wStr)
                .TextMatrix(ID, 12) = Format(ZSR.合計重量, "#########0.00")
                .TextMatrix(ID, 13) = Format(ZSR.材料管理番号, "##########")
                .TextMatrix(ID, 14) = Format(ZSR.板厚, "####.00")
                .TextMatrix(ID, 15) = Format(ZSR.幅, "####.00")
                .TextMatrix(ID, 16) = Format(ZSR.長さ, "####.00")
                .TextMatrix(ID, 17) = Format(ZSR.ラベル重量1, "########0.000")
                .TextMatrix(ID, 18) = Format(ZSR.ラベル本数1, "#0")
                .TextMatrix(ID, 19) = Format(ZSR.ラベル重量2, "########0.000")
                .TextMatrix(ID, 20) = Format(ZSR.ラベル本数2, "#0")
                .TextMatrix(ID, 21) = Format(ZSR.ラベル重量3, "########0.000")
                .TextMatrix(ID, 22) = Format(ZSR.ラベル本数3, "#0")
                .TextMatrix(ID, 23) = Format(ZSR.ラベル重量4, "########0.000")
                .TextMatrix(ID, 24) = Format(ZSR.ラベル本数4, "#0")
                .TextMatrix(ID, 25) = Format(ZSR.戻し重量, "#")
                .TextMatrix(ID, 26) = Format(ZSR.材料使用量管理NO, "####")
                .TextMatrix(ID, 27) = Format(ZSR.材料使用量管理行, "####")
                .TextMatrix(ID, 28) = ZSR.作業日
                .TextMatrix(ID, 29) = ZSR.部門
                .TextMatrix(ID, 30) = ZSR.作業者
                .TextMatrix(ID, 31) = Format("")
                .TextMatrix(ID, 32) = Format("")
                .TextMatrix(ID, 33) = Format("")
                
                .TextMatrix(ID, 34) = RTrim(ZSR.備考)
            End With
            ID = ID + 1
            Call ZSRReadNext
        Loop
        With vsGrid
             .Row = 1: .Col = 1: .ColSel = .Cols - 1
             .Redraw = True
        End With

        GridSet = True
    End If
    
    Call ZSRClose
End Function
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
'-------------------------------------------------------------------' 2009.12.14 del start (仮)
'    Input_Area.Visible = EnabledFlg
'    B_TabDummy.Visible = EnabledFlg
'-------------------------------------------------------------------' 2009.12.14 del end   (仮)
'   *** ↑: 上記の１行は、入力処理時にグリッド移動が発生しないようにする為！ ***
End Sub

'-----------------------------------------------------------------------------------------------' 2009.12.16 del start
''==========================================================='
''   グリッド移動
''==========================================================='
'Private Sub Grid_RowChange(KeyCode As Integer)
'    If KeyCode <> vbKeyUp And _
'       KeyCode <> vbKeyDown And _
'       KeyCode <> vbKeyPageUp And _
'       KeyCode <> vbKeyPageDown Then Exit Sub
'
''   テキストからグリッドに退避
'    Call TextToVsGrid(Val(Input_Area.Tag))
'    gInt = vsGrid.Row
'   '↑
'    If KeyCode = vbKeyUp Then
'        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
'        gInt = gInt - 1
'    End If
'   '↓
'    If KeyCode = vbKeyDown Then
'        If Not wAllAdd Then
'           If gInt >= vsGrid.Rows - 1 Then GoTo Grid_RowChange_End
'           If RTrim(vsGrid.TextMatrix(gInt + 1, 0)) = "" Then GoTo Grid_RowChange_End
'        End If
'        If gInt + 1 > vsGrid.Rows - 1 Then
'            If RTrim(vsGrid.TextMatrix(gInt, 1)) <> "" Then
'                vsGrid.Rows = gInt + 1 + 1
''                vsGrid.TextMatrix(gInt + 1, 0) = gInt + 1
'            Else
'                GoTo Grid_RowChange_End
'            End If
'        End If
'        gInt = gInt + 1
'    End If
'   'PageUp
'    If KeyCode = vbKeyPageUp Then
'        If gInt - CisVsGrid.DispMax < 1 Then
'            gInt = 1
'        Else
'            gInt = gInt - CisVsGrid.DispMax
'        End If
'    End If
'   'PageDown
'    If KeyCode = vbKeyPageDown Then
'        If gInt + CisVsGrid.DispMax > vsGrid.Rows - 1 Then
'            gInt = vsGrid.Rows - 1
'        Else
'            gInt = gInt + CisVsGrid.DispMax
'        End If
'    End If
'    vsGrid.Row = gInt
''   グリッドからテキストに表示
'    Call VsGridToText(vsGrid.Row)
'Grid_RowChange_End:
'    KeyCode = 0
'    Me.Refresh
'    DoEvents
'End Sub
'-----------------------------------------------------------------------------------------------' 2009.12.16 del end
'==========================================================='
'   グリッドから入力テキストに移す
'==========================================================='
Private Sub VsGridToText(RowNo As Integer)
    With vsGrid
        If Trim(.TextMatrix(RowNo, 1)) = "" Then Exit Sub
        '材料管理番号
        mZRKanriNo = CisFun.Val2(.TextMatrix(RowNo, 13))
        '品番
        iB1_Hinbn = RTrim(.TextMatrix(RowNo, 1))
        '材質
        iB1_Zaist = RTrim(.TextMatrix(RowNo, 2))
        '板厚
        iB1_Itatu = RTrim(.TextMatrix(RowNo, 14))
        '幅
        iB1_Width = RTrim(.TextMatrix(RowNo, 15))
        '長さ
        iB1_Long = Format(CisFun.Val2(.TextMatrix(RowNo, 16)), "####")
        '加工数
        iB1_Kakou = Format(CisFun.Val2(.TextMatrix(RowNo, 4)), "####")
        '単重
        B1lb_TanJu = Format(CisFun.Val2(.TextMatrix(RowNo, 5)), "###0.000")
        '使用量
        iB1_Shiyo = Format(CisFun.Val2(.TextMatrix(RowNo, 6)), "######")
        'ラベル重量・本数 1
        iB1_LblJ1 = Format(CisFun.Val2(.TextMatrix(RowNo, 17)), "###0.000")
        iB1_LblH1 = Format(CisFun.Val2(.TextMatrix(RowNo, 18)), "##")
        'ラベル重量・本数 2
        iB1_LblJ2 = Format(CisFun.Val2(.TextMatrix(RowNo, 19)), "###0.000")
        iB1_LblH2 = Format(CisFun.Val2(.TextMatrix(RowNo, 20)), "##")
        'ラベル重量・本数 3
        iB1_LblJ3 = Format(CisFun.Val2(.TextMatrix(RowNo, 21)), "###0.000")
        iB1_LblH3 = Format(CisFun.Val2(.TextMatrix(RowNo, 22)), "##")
        'ラベル重量・本数 4
        iB1_LblJ4 = Format(CisFun.Val2(.TextMatrix(RowNo, 23)), "###0.000")
        iB1_LblH4 = Format(CisFun.Val2(.TextMatrix(RowNo, 24)), "##")
        ' 戻し重量
        Chk_Modosi.Value = CisFun.Val2(.TextMatrix(RowNo, 25))
        ' 合計重量
        iB1_Gokei = .TextMatrix(RowNo, 12)
        
        ' 備考
        iB1_Biko = .TextMatrix(RowNo, 34)
        ' 在庫数
        Call DispZaikoSu        ' 2009.12.18 add
    End With

End Sub
'+---------------------------------------+
'+  グリッド (KeyDown)
'+---------------------------------------+
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And _
       KeyCode <> vbKeyDown And _
       KeyCode <> vbKeyPageUp And _
       KeyCode <> vbKeyPageDown Then Exit Sub
   'Call TextToVsGrid(Val(Input_Area.Tag))
    Call HeadBodyClear("B1")
    Call VsGridToText(vsGrid.Row)
End Sub
'+---------------------------------------+
'+  グリッド (Click)
'+---------------------------------------+
Private Sub VSGrid_Click()
    With vsGrid
'        If ProcHB = "B1" Then
'            .HighLight = flexHighlightNever
'        Else
'            .HighLight = flexHighlightWithFocus
'        End If
        
        If mSelectGyo = 0 And SyoriKB <> "ADD" Then     ' If ProcHB <> "B1" Then    ' If SyoriKB <> "ADD" Then  ' If ProcHB = "B2" Then
            Call HeadBodyClear("B1")

            Call VsGridToText(vsGrid.Row)
            Call vsGrid.SetFocus
        Else
            If SyoriKB <> "DEL" Then ActObj.SetFocus
        End If
    End With
End Sub
'+---------------------------------------+
'+  グリッド (DoubleClick)
'+---------------------------------------+
Private Sub VSGrid_DblClick()
    With vsGrid
        If .TextMatrix(.Row, 0) = "" Or SyoriKB = "ADD" Then
           If Mid(ActObj.Name, 1, 2) = "iH" Then Set ActObj = iB1_Hinbn
           If B1_Area1.Enabled = True Then ActObj.SetFocus
           Exit Sub
        End If
        
        If Trim(.TextMatrix(.Row, 1)) = "" Then Exit Sub
        
        Call GridNoSelect           ' 現在選択されている行があれば非選択にする
        Call CisVsGrid.vsColor
        Call VsGridToText(.Row)
        mSelectGyo = .Row
        
        If SyoriKB = "MNT" Then
            ProcHB = "B1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
            iB1_Hinbn.SetFocus
            B1_Area2.Enabled = True
            B1_Area2.BackColor = gPGotSel
            PB_ENT.Visible = True
        End If

        If SyoriKB = "DEL" Then
            PB_ENT.Visible = True
        End If
    End With
    
'---------------------------------------------------------------------------' 2009.12.14 del start
'    If RTrim(vsGrid.TextMatrix(vsGrid.Row, 1)) = "" Then
'        Exit Sub
'    End If
'    With CisFun
'        .MB_Lines = 4
'        .MB_MSG(2) = "    " & vsGrid.Row & " 行目をクリアしますか？            "
'        .MB_Title = ""
'        .MB_Button = OK_CAN
'        If Not .MBOX Then Exit Sub
'    End With
'    iB1_KakoCD = ""
'    B1lb_Kakonm = ""
'    iB1_KikaiCD = ""
'    B1lb_Kikainm = ""
'    iB1_Start = ""
'    iB1_End = ""
'    iB1_Kanse = ""
'    iB1_Furyo = ""
'    B1lb_TimeSu = ""
'    Call TextToVsGrid(Val(Input_Area.Tag))
'---------------------------------------------------------------------------' 2009.12.14 del end
End Sub
Private Sub vsGrid_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
End Sub
'+-----------------------+
'+  グリッド内容初期化    +
'+-----------------------+
Private Sub GridClear()
    Dim Cnt1    As Integer
    Dim Cnt2    As Integer
    Dummy.Enabled = True: Dummy.SetFocus
    
    With vsGrid
        .Enabled = False
        .Rows = 11
        For Cnt1 = 1 To .Rows - 1
            For Cnt2 = 0 To .Cols - 1
                .TextMatrix(Cnt1, Cnt2) = ""
            Next Cnt2
        Next Cnt1
        .Enabled = True
    End With
   'Call GridInit
    
    Dummy.Enabled = False
End Sub
'+-----------------------+
'+  グリッド選択行解除
'+-----------------------+
Private Sub GridNoSelect()
    Dim wCnt1   As Integer
    Dim wRow1   As Integer
    Dim wB2ED   As Boolean
    
    With vsGrid
        wRow1 = .Row
        wB2ED = B1_Area2.Enabled
        B1_Area2.Enabled = True
        For wCnt1 = 1 To .Rows - 1
            If .TextMatrix(wCnt1, CisVsGrid.SelectCol) <> "" Then
                .Row = wCnt1
                Call CisVsGrid.vsColor
            End If
        Next
        .Row = wRow1
        B1_Area2.Enabled = wB2ED
    End With
End Sub

Private Function Get_Zairyo()
    Get_Zairyo = False
    
'   << 材料マスタ >>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  zrm.材質"
    gSL_Select = gSL_Select & ", zrm.板厚"
    gSL_Select = gSL_Select & ", zrm.幅"
    gSL_Select = gSL_Select & ", zrm.長さ"
    gSL_Select = gSL_Select & ", zhm.使用量"
    gSL_Select = gSL_Select & ", zrm.材料管理番号 "
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  材料品番マスタ zhm"
    gSL_Select = gSL_Select & ", 材料マスタ zrm"
    gSL_Select = gSL_Select & " where zhm.使用品番 = '" & RTrim(iB1_Hinbn.Text) & "'"
    gSL_Select = gSL_Select & "   and zrm.材料管理番号 = zhm.材料管理番号"
    gSL_Select = gSL_Select & "   and zrm.形状 = '1'"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  zrm.材料管理番号"
    If Not ZHMRead(gSL_Select, , 1) Then
        Call ItemsClearZHM
        Exit Function
    End If

    ZRM.材料管理番号 = ZHM.材料管理番号
    Call ZRMRead("", 1)
    
    Get_Zairyo = True
End Function
'-------------------------------------------'
'   材料使用量管理Noの獲得                            '
'-------------------------------------------'
Private Function GetKanriNo() As Boolean
    GetKanriNo = False
    
    With CisDB
        .SQL = "番号獲得"
        .StoadoCount = 3
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = OutPut用
        .ParaNo = 3: .ParaIO = Return用
    '
        .ParaNo = 1: .ParaValue = 22
    '
        If Not .DBStored Then
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "    材料使用量管理NOの獲得に失敗しました            "
            CisFun.MB_Title = "採番処理"
            CisFun.MB_Button = Error
            CisFun.MBOX
            Exit Function
        Else
            .ParaNo = 2: mZSKanriNo = .ParaValue
        End If
    End With

    GetKanriNo = True
End Function

'---------------------------------------------------------------------------' 2009.12.16 del start
''+---------------------------------------+
''+      名称マスタ初期設定
''+---------------------------------------+
'Private Sub MeisyoInit()
'' 停止名（数字１＝　0:トラブル 1:段取り 2:計画停止 3:試作）
'    Call ItemsClearSYMCreate
'    With SYMCreate
'        .区分名称 = "停止名"
'        .区分桁数 = 2
'        .制御区分 = ""
'        .作成者 = gTanto
'        .検索名 = "停止名　検索"
'        .コード名 = "停止名"
'        .名称名 = "停止名"
'    End With
'    If Not SYMInitCreate(True) Then
'        With SYMCreate
'            .値 = "1"
'            .値名称 = "金型取付"
'            .数字1 = 1
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "2"
'            .値名称 = "金型取外し"
'            .数字1 = 1
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "3"
'            .値名称 = "寸法チェック"
'            .数字1 = 1
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "4"
'            .値名称 = "初物チェック"
'            .数字1 = 1
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "5"
'            .値名称 = "終物チェック"
'            .数字1 = 1
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "6"
'            .値名称 = "材料運搬"
'            .数字1 = 1
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "7"
'            .値名称 = "製品運搬"
'            .数字1 = 1
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "8"
'            .値名称 = "片付け"
'            .数字1 = 1
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "9"
'            .値名称 = "通箱運搬"
'            .数字1 = 1
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "14"
'            .値名称 = "コイル巻き戻し"
'            .数字1 = 1
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "23"
'            .値名称 = "ティーチング"
'            .数字1 = 0
'        End With
'        Call SYMInitCreate
'
'        With SYMCreate
'            .値 = "10"
'            .値名称 = "機械トラブル"
'            .数字1 = 0
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "11"
'            .値名称 = "材料トラブル"
'            .数字1 = 0
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "12"
'            .値名称 = "品質トラブル"
'            .数字1 = 0
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "13"
'            .値名称 = "金型トラブル"
'            .数字1 = 0
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "15"
'            .値名称 = "クレーン待ち"
'            .数字1 = 0
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "16"
'            .値名称 = "リフト待ち"
'            .数字1 = 0
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "17"
'            .値名称 = "工程待ち"
'            .数字1 = 0
'        End With
'        Call SYMInitCreate
'
'        With SYMCreate
'            .値 = "18"
'            .値名称 = "朝礼５Ｓ"
'            .数字1 = 2
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "19"
'            .値名称 = "応援"
'            .数字1 = 2
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "20"
'            .値名称 = "計画停止"
'            .数字1 = 2
'        End With
'        Call SYMInitCreate
'        With SYMCreate
'            .値 = "21"
'            .値名称 = "その他ウール詰"
'            .数字1 = 2
'        End With
'        Call SYMInitCreate
'
'        With SYMCreate
'            .値 = "22"
'            .値名称 = "試作"
'            .数字1 = 3
'        End With
'        Call SYMInitCreate
'    End If
'
'End Sub
'---------------------------------------------------------------------------' 2009.12.16 del end

'-------------------------------------------'
'   在庫数表示メインモジュール
'-------------------------------------------'
Private Function DispZaikoSu() As Boolean
    If mDebugFlg Then L_TestDisp = ""
    DispZaikoSu = False
    
    Dim wStr    As String
    
    B1lb_Zaiko = ""
   'Call GetTanju           ' 材料管理番号再取得　：ここに入る前にセットされているので不要
    
    If mZRKanriNo = 0 Then
        B1lb_Zaiko = ""     ' 材料管理番号が取得できなければ在庫なしで終了
        Exit Function
    End If
    
    gSL_Select = ""                                             ' 最新在庫年月取得
    gSL_Select = gSL_Select & "SELECT TOP(1) * FROM 材料在庫マスタ"
    gSL_Select = gSL_Select & " WHERE 材料管理番号 = " & Str(mZRKanriNo)
    gSL_Select = gSL_Select & " ORDER BY 年月度 DESC"
    
    If Not ZZMRead(gSL_Select) Then
' 在庫マスタ[無]のとき
        Call ZaikoKeisan1
        GoTo DispZaikoSu_ED
    End If
    
' 在庫マスタ[有]実棚日[空白]のとき
    Call ZaikoKeisan2
    GoTo DispZaikoSu_ED
    
DispZaikoSu_ED:
    wStr = ""
    If mZaikoSu <> 0 Then wStr = Format(mZaikoSu, "###0.000")
   'If mDebugFlg Then wStr = Format(mZaikoSu, "###0.000")
    B1lb_Zaiko = wStr
    DispZaikoSu = True
End Function

Private Function ZaikoKeisan1() As Boolean
    ZaikoKeisan1 = False
        
    Dim wSuIn   As Double
    Dim wSuOut  As Double
    wSuIn = 0:  wSuOut = 0
    
    On Error GoTo ZaikoKeisan1_Err
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT TOP(1)"
    gSL_Select = gSL_Select & "       ISNULL( ZNT.入荷数, 0 ) + ISNULL( ZZT.実績数, 0 ) AS 入荷数"
    gSL_Select = gSL_Select & " FROM  材料入荷テーブル ZN"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
    gSL_Select = gSL_Select & "           (SELECT NK.材料管理番号, SUM(NK.入荷数) AS 入荷数 FROM 材料入荷テーブル NK"
    gSL_Select = gSL_Select & "             WHERE SUBSTRING(NK.入荷日, 1, 6) >= '" & Mid(iH1_SYMD, 1, 6) & "'"
    gSL_Select = gSL_Select & "               AND NK.材料管理番号 =  " & mZRKanriNo
    gSL_Select = gSL_Select & "             GROUP BY NK.材料管理番号) ZNT"
    gSL_Select = gSL_Select & "         ON ZNT.材料管理番号 = ZN.材料管理番号"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
    gSL_Select = gSL_Select & "           (SELECT ZZ.実績区分, SUM(ZZ.実績数) AS 実績数 FROM 材料実績管理テーブル ZZ"
    gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZZ.実績日, 1,6) >= '" & Mid(iH1_SYMD, 1, 6) & "'"
    gSL_Select = gSL_Select & "               AND ZZ.材質     = '" & Trim(iB1_Zaist) & "'"
    gSL_Select = gSL_Select & "               AND ZZ.板厚     =  " & CisFun.Val2(iB1_Itatu)
    gSL_Select = gSL_Select & "               AND ZZ.幅       =  " & CisFun.Val2(iB1_Width)
    gSL_Select = gSL_Select & "               AND ZZ.長さ     =  " & CisFun.Val2(iB1_Long)
''  gSL_Select = gSL_Select & "               AND ZZ.実績区分 =  2"     ' 2010.01.22 del
    gSL_Select = gSL_Select & "               AND ZZ.実績区分 = 1"      ' 2010.01.22 add
    gSL_Select = gSL_Select & "             GROUP BY ZZ.実績区分) ZZT"
''  gSL_Select = gSL_Select & "         ON ZZT.実績区分 = 2"            ' 2010.01.18 del
    gSL_Select = gSL_Select & "         ON ZZT.実績区分 = 1"            ' 2010.01.18 add
    gSL_Select = gSL_Select & " WHERE SUBSTRING(ZN.入荷日, 1, 6) >= '" & Mid(iH1_SYMD, 1, 6) & "'"
    gSL_Select = gSL_Select & "   AND ZN.材料管理番号 =  " & mZRKanriNo
    If ZYTRead(gSL_Select) Then wSuIn = ZYT.入荷数
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT TOP(1)"
    gSL_Select = gSL_Select & "       ISNULL( ZST.入荷数, 0 ) + ISNULL( ZZT.実績数, 0 ) AS 使用量"
    gSL_Select = gSL_Select & " FROM  材料使用量テーブル ZS"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
    gSL_Select = gSL_Select & "           (SELECT SR.材料管理番号, SUM(SR.使用量) AS 入荷数 FROM 材料使用量テーブル SR"
    gSL_Select = gSL_Select & "             WHERE SUBSTRING(SR.作業日, 1, 6) >= '" & Mid(iH1_SYMD, 1, 6) & "'"
    gSL_Select = gSL_Select & "               AND SR.材料管理番号 =  " & mZRKanriNo
    gSL_Select = gSL_Select & "             GROUP BY SR.材料管理番号) ZST"
    gSL_Select = gSL_Select & "         ON ZST.材料管理番号 = ZS.材料管理番号"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
    gSL_Select = gSL_Select & "           (SELECT ZZ.実績区分,ZZ.入力種類, SUM(ZZ.実績数) AS 実績数 FROM 材料実績管理テーブル ZZ"   ' 2010.01.18 upd : 入力種類を追加
    gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZZ.実績日, 1, 6) >= '" & Mid(iH1_SYMD, 1, 6) & "'"
    gSL_Select = gSL_Select & "               AND ZZ.材質     = '" & Trim(iB1_Zaist) & "'"
    gSL_Select = gSL_Select & "               AND ZZ.板厚     =  " & CisFun.Val2(iB1_Itatu)
    gSL_Select = gSL_Select & "               AND ZZ.幅       =  " & CisFun.Val2(iB1_Width)
    gSL_Select = gSL_Select & "               AND ZZ.長さ     =  " & CisFun.Val2(iB1_Long)
''  gSL_Select = gSL_Select & "               AND ZZ.実績区分 =  3"     ' 2010.01.18 del
    gSL_Select = gSL_Select & "               AND ZZ.実績区分 <> 1"     ' 2010.01.18 add
    gSL_Select = gSL_Select & "               AND ZZ.入力種類 <> 'A'"   ' 2010.01.18 add
    gSL_Select = gSL_Select & "             GROUP BY ZZ.実績区分,ZZ.入力種類) ZZT"      ' 2010.01.18 upd : 入力種類を追加
''  gSL_Select = gSL_Select & "         ON ZZT.実績区分 = 3"            ' 2010.01.18 del
    gSL_Select = gSL_Select & "         ON ZZT.実績区分 <> 1"           ' 2010.01.18 add
    gSL_Select = gSL_Select & "        AND ZZT.入力種類 <> 'A'"         ' 2010.01.18 add
    gSL_Select = gSL_Select & " WHERE SUBSTRING(ZS.作業日, 1, 6) >= '" & Mid(iH1_SYMD, 1, 6) & "'"
    gSL_Select = gSL_Select & "   AND ZS.材料管理番号 =  " & mZRKanriNo
    
    If ZSRRead(gSL_Select) Then wSuOut = ZSR.使用量
    
    On Error GoTo 0
    mZaikoSu = wSuIn - wSuOut
    If mDebugFlg Then L_TestDisp.Caption = "入庫 = " & Format(wSuIn, "00000000.000") _
                                & vbCrLf & "出庫 = " & Format(wSuOut, "00000000.000")
    ZaikoKeisan1 = True
    Exit Function

ZaikoKeisan1_Err:
    Exit Function
End Function
'
'
Private Function ZaikoKeisan2() As Boolean
    ZaikoKeisan2 = False
    
    Dim wNendo  As String   ' 年月度
    Dim wJTDay  As String   ' 実棚日
    Dim wSuIn   As Double
    Dim wSuOut  As Double
    wSuIn = 0:  wSuOut = 0
    
    On Error GoTo ZaikoKeisan2_Err

    wNendo = ZZM.年月度
    wJTDay = ZZM.実棚日
    
    If Trim(ZZM.実棚日) = "" Then
    ' 実棚日が空白の場合の処理
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT TOP(1)"
        gSL_Select = gSL_Select & "       ZM.在庫数 AS 在マ在庫数,ISNULL( ZNT.入荷数, 0 ) AS 入荷数, ISNULL( ZZT.実績数, 0 ) AS 実績数,"    ' この行は試験確認用
        gSL_Select = gSL_Select & "       ZM.在庫数 + ISNULL( ZNT.入荷数, 0 ) + ISNULL( ZZT.実績数, 0 ) AS 在庫数 FROM  材料在庫マスタ ZM"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
        gSL_Select = gSL_Select & "           (SELECT NK.材料管理番号, SUM(NK.入荷数) AS 入荷数 FROM 材料入荷テーブル NK"
        gSL_Select = gSL_Select & "             WHERE SUBSTRING(NK.入荷日, 1, 6) >= '" & wNendo & "'"
        gSL_Select = gSL_Select & "               AND NK.材料管理番号 =  " & mZRKanriNo
        gSL_Select = gSL_Select & "             GROUP BY NK.材料管理番号) ZNT"
        gSL_Select = gSL_Select & "         ON ZNT.材料管理番号 = ZM.材料管理番号"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
        gSL_Select = gSL_Select & "           (SELECT ZZ.実績区分, SUM(ZZ.実績数) AS 実績数 FROM 材料実績管理テーブル ZZ"
        gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZZ.実績日, 1, 6) >= '" & wNendo & "'"
        gSL_Select = gSL_Select & "               AND ZZ.材質     = '" & Trim(iB1_Zaist) & "'"
        gSL_Select = gSL_Select & "               AND ZZ.板厚     =  " & CisFun.Val2(iB1_Itatu)
        gSL_Select = gSL_Select & "               AND ZZ.幅       =  " & CisFun.Val2(iB1_Width)
        gSL_Select = gSL_Select & "               AND ZZ.長さ     =  " & CisFun.Val2(iB1_Long)
''      gSL_Select = gSL_Select & "               AND ZZ.実績区分 = 2"      ' 2010.01.22 del
        gSL_Select = gSL_Select & "               AND ZZ.実績区分 = 1"      ' 2010.01.22 add
        gSL_Select = gSL_Select & "             GROUP BY ZZ.実績区分) ZZT"
''      gSL_Select = gSL_Select & "         ON ZZT.実績区分 = 2"            ' 2010.01.18 del
        gSL_Select = gSL_Select & "         ON ZZT.実績区分 = 1"            ' 2010.01.18 add
        gSL_Select = gSL_Select & " WHERE ZM.年月度      >= '" & wNendo & "'"
        gSL_Select = gSL_Select & "   AND ZM.材料管理番号 =  " & mZRKanriNo
        If ZZMRead(gSL_Select) Then wSuIn = ZZM.在庫数

        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT TOP(1)"
        gSL_Select = gSL_Select & "       ISNULL( ZST.入荷数, 0 ) + ISNULL( ZZT.実績数, 0 ) AS 使用量"
        gSL_Select = gSL_Select & " FROM  材料使用量テーブル ZS"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
        gSL_Select = gSL_Select & "           (SELECT SR.材料管理番号, SUM(SR.使用量) AS 入荷数 FROM 材料使用量テーブル SR"
        gSL_Select = gSL_Select & "             WHERE SUBSTRING(SR.作業日, 1, 6) >= '" & wNendo & "'"
        gSL_Select = gSL_Select & "               AND SR.材料管理番号 =  " & mZRKanriNo
        gSL_Select = gSL_Select & "             GROUP BY SR.材料管理番号) ZST"
        gSL_Select = gSL_Select & "         ON ZST.材料管理番号 = ZS.材料管理番号"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
        gSL_Select = gSL_Select & "           (SELECT ZZ.実績区分,ZZ.入力種類, SUM(ZZ.実績数) AS 実績数 FROM 材料実績管理テーブル ZZ"   ' 2010.01.18 upd : 入力種類を追加
        gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZZ.実績日, 1, 6) >= '" & wNendo & "'"
        gSL_Select = gSL_Select & "               AND ZZ.材質     = '" & Trim(iB1_Zaist) & "'"
        gSL_Select = gSL_Select & "               AND ZZ.板厚     =  " & CisFun.Val2(iB1_Itatu)
        gSL_Select = gSL_Select & "               AND ZZ.幅       =  " & CisFun.Val2(iB1_Width)
        gSL_Select = gSL_Select & "               AND ZZ.長さ     =  " & CisFun.Val2(iB1_Long)
''      gSL_Select = gSL_Select & "               AND ZZ.実績区分 =  3"     ' 2010.01.18 del
        gSL_Select = gSL_Select & "               AND ZZ.実績区分 <> 1"     ' 2010.01.18 add
        gSL_Select = gSL_Select & "               AND ZZ.入力種類 <> 'A'"   ' 2010.01.18 add
        gSL_Select = gSL_Select & "             GROUP BY ZZ.実績区分,ZZ.入力種類) ZZT"      ' 2010.01.18 upd : 入力種類を追加
''      gSL_Select = gSL_Select & "         ON ZZT.実績区分 = 3"            ' 2010.01.18 del
        gSL_Select = gSL_Select & "         ON ZZT.実績区分 <> 1"           ' 2010.10.18 add
        gSL_Select = gSL_Select & "        AND ZZT.入力種類 <> 'A'"         ' 2010.01.18 add
        gSL_Select = gSL_Select & " WHERE SUBSTRING(ZS.作業日, 1, 6) >= '" & wNendo & "'"
        gSL_Select = gSL_Select & "   AND ZS.材料管理番号 =  " & mZRKanriNo
        If ZSRRead(gSL_Select) Then wSuOut = ZSR.使用量
    Else
    ' 実棚日が入力済の場合の処理
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT TOP(1)"
        gSL_Select = gSL_Select & "       ZM.在庫数 AS 在マ在庫数,ISNULL( ZNT.入荷数, 0 ) AS 入荷数, ISNULL( ZZT.実績数, 0 ) AS 実績数,"    ' この行は試験確認用
        gSL_Select = gSL_Select & "       ZM.在庫数 + ISNULL( ZNT.入荷数, 0 ) + ISNULL( ZZT.実績数, 0 ) AS 在庫数 FROM  材料在庫マスタ ZM"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
        gSL_Select = gSL_Select & "           (SELECT NK.材料管理番号, SUM(NK.入荷数) AS 入荷数 FROM 材料入荷テーブル NK"
        gSL_Select = gSL_Select & "             WHERE NK.入荷日       > '" & wJTDay & "'"
        gSL_Select = gSL_Select & "               AND NK.材料管理番号 =  " & mZRKanriNo
        gSL_Select = gSL_Select & "             GROUP BY NK.材料管理番号) ZNT"
        gSL_Select = gSL_Select & "         ON ZNT.材料管理番号 = ZM.材料管理番号"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
        gSL_Select = gSL_Select & "           (SELECT ZZ.実績区分, SUM(ZZ.実績数) AS 実績数 FROM 材料実績管理テーブル ZZ"
        gSL_Select = gSL_Select & "             WHERE ZZ.実績日   > '" & wJTDay & "'"
        gSL_Select = gSL_Select & "               AND ZZ.材質     = '" & Trim(iB1_Zaist) & "'"
        gSL_Select = gSL_Select & "               AND ZZ.板厚     =  " & CisFun.Val2(iB1_Itatu)
        gSL_Select = gSL_Select & "               AND ZZ.幅       =  " & CisFun.Val2(iB1_Width)
        gSL_Select = gSL_Select & "               AND ZZ.長さ     =  " & CisFun.Val2(iB1_Long)
        gSL_Select = gSL_Select & "               AND ZZ.実績区分 = 2"      ' 2010.01.22 del
        gSL_Select = gSL_Select & "               AND ZZ.実績区分 = 2"      ' 2010.01.22 add
        gSL_Select = gSL_Select & "             GROUP BY ZZ.実績区分) ZZT"
''      gSL_Select = gSL_Select & "         ON ZZT.実績区分 = 2"            ' 2010.01.18 del
        gSL_Select = gSL_Select & "         ON ZZT.実績区分 = 1"            ' 2010.01.18 add
        gSL_Select = gSL_Select & " WHERE ZM.年月度      >= '" & Mid(wJTDay, 1, 6) & "'"
        gSL_Select = gSL_Select & "   AND ZM.材料管理番号 =  " & mZRKanriNo
        If ZZMRead(gSL_Select) Then wSuIn = ZZM.在庫数

        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT TOP(1)"
        gSL_Select = gSL_Select & "       ISNULL( ZST.入荷数, 0 ) + ISNULL( ZZT.実績数, 0 ) AS 使用量"
        gSL_Select = gSL_Select & " FROM  材料使用量テーブル ZS"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
        gSL_Select = gSL_Select & "           (SELECT SR.材料管理番号, SUM(SR.使用量) AS 入荷数 FROM 材料使用量テーブル SR"
        gSL_Select = gSL_Select & "             WHERE SR.作業日       > '" & wJTDay & "'"
        gSL_Select = gSL_Select & "               AND SR.材料管理番号 =  " & mZRKanriNo
        gSL_Select = gSL_Select & "             GROUP BY SR.材料管理番号) ZST"
        gSL_Select = gSL_Select & "         ON ZST.材料管理番号 = ZS.材料管理番号"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
        gSL_Select = gSL_Select & "           (SELECT ZZ.実績区分, SUM(ZZ.実績数) AS 実績数, ZZ.入力種類 FROM 材料実績管理テーブル ZZ"
        gSL_Select = gSL_Select & "             WHERE ZZ.実績日   > '" & wJTDay & "'"
        gSL_Select = gSL_Select & "               AND ZZ.材質     = '" & Trim(iB1_Zaist) & "'"
        gSL_Select = gSL_Select & "               AND ZZ.板厚     =  " & CisFun.Val2(iB1_Itatu)
        gSL_Select = gSL_Select & "               AND ZZ.幅       =  " & CisFun.Val2(iB1_Width)
        gSL_Select = gSL_Select & "               AND ZZ.長さ     =  " & CisFun.Val2(iB1_Long)
''      gSL_Select = gSL_Select & "               AND ZZ.実績区分 =  3"     ' 2010.01.18 del
        gSL_Select = gSL_Select & "               AND ZZ.実績区分 <> 1"     ' 2010.01.18 add
        gSL_Select = gSL_Select & "               AND ZZ.入力種類 <> 'A'"   ' 2010.01.18 add
        gSL_Select = gSL_Select & "             GROUP BY ZZ.実績区分, ZZ.入力種類) ZZT"
''      gSL_Select = gSL_Select & "         ON ZZT.実績区分 = 3"            ' 2010.01.18 del
        gSL_Select = gSL_Select & "        ON  ZZT.実績区分 <> 1"           ' 2010.01.18 add
        gSL_Select = gSL_Select & "        AND ZZT.入力種類 <> 'A'"         ' 2010.01.18 add
        gSL_Select = gSL_Select & " WHERE ZS.作業日       > '" & wJTDay & "'"
        gSL_Select = gSL_Select & "   AND ZS.材料管理番号 =  " & mZRKanriNo
        If ZSRRead(gSL_Select) Then wSuOut = ZSR.使用量
    End If

    On Error GoTo 0
    mZaikoSu = wSuIn - wSuOut
    If mDebugFlg Then L_TestDisp.Caption = "入庫 = " & Format(wSuIn, "00000000.000") _
                                & vbCrLf & "出庫 = " & Format(wSuOut, "00000000.000")
    ZaikoKeisan2 = True
    Exit Function

ZaikoKeisan2_Err:
    Exit Function
End Function



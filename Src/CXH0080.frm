VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXH0080 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "発注指示(発注点）"
   ClientHeight    =   11235
   ClientLeft      =   720
   ClientTop       =   1185
   ClientWidth     =   15360
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D Message 
      Height          =   930
      Left            =   2385
      Top             =   4935
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
      Begin Cis3D_v60.CIS3D MessageDisp 
         Height          =   525
         Left            =   240
         Top             =   210
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   926
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "【表示中】しばらくお待ち下さい"
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
   Begin Cis3D_v60.CIS3D B1_Area1 
      Height          =   8355
      Left            =   120
      Top             =   2325
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   14737
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
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   8175
         Left            =   60
         Top             =   90
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   14420
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         Begin vsFlexLib.vsFlexArray vsGrid1 
            DragIcon        =   "CXH0080.frx":0000
            Height          =   8040
            Left            =   60
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   60
            Width           =   9270
            _Version        =   196608
            _ExtentX        =   16351
            _ExtentY        =   14182
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
            FormatString    =   $"CXH0080.frx":030A
            Rows            =   23
            Cols            =   34
            BackColor       =   16777215
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   2
            RowHeightMin    =   350
            GridLines       =   2
            MergeCells      =   3
            BackColorAlternate=   16777215
            Begin CisYMD_With_Btn.CisYMDwB iB1_NYmd 
               Height          =   360
               Left            =   4530
               TabIndex        =   7
               Top             =   330
               Visible         =   0   'False
               Width           =   1425
               _ExtentX        =   2249
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
               Object.Width           =   1425
               cSize           =   1
               cChkResult      =   0   'False
               cType           =   1
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
            Begin CisText_V60.CisText iB1_HSuryo 
               Height          =   360
               Left            =   3240
               TabIndex        =   6
               Top             =   330
               Visible         =   0   'False
               Width           =   1320
               _ExtentX        =   2328
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
         Begin Cis3D_v60.CIS3D B2_Area1 
            Height          =   8040
            Left            =   9360
            Top             =   60
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   14182
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   7860
               Left            =   90
               Top             =   90
               Width           =   5415
               _ExtentX        =   9551
               _ExtentY        =   13864
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
               Begin Cis3D_v60.CIS3D CIS3D5 
                  Height          =   405
                  Left            =   30
                  Top             =   3480
                  Width           =   5325
                  _ExtentX        =   9393
                  _ExtentY        =   714
                  BackColor       =   16761024
                  Caption         =   "【　発 注 先 情 報  】"
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
                  cAlingnment     =   7
               End
               Begin vsFlexLib.vsFlexArray VsGridH 
                  Height          =   1410
                  Left            =   30
                  TabIndex        =   9
                  TabStop         =   0   'False
                  Top             =   3870
                  Width           =   5310
                  _Version        =   196608
                  _ExtentX        =   9366
                  _ExtentY        =   2487
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
                  FormatString    =   "No|仕入先|受入|発注率|"
                  Rows            =   4
                  Cols            =   6
                  BackColor       =   16761024
                  BackColorBkg    =   14737632
                  FocusRect       =   0
                  HighLight       =   2
                  ScrollBars      =   2
                  RowHeightMin    =   330
                  GridColor       =   8421504
                  SelectionMode   =   1
                  BackColorAlternate=   16761024
               End
               Begin vsFlexLib.vsFlexArray VsGridY 
                  Height          =   2070
                  Left            =   30
                  TabIndex        =   10
                  TabStop         =   0   'False
                  Top             =   5670
                  Width           =   5310
                  _Version        =   196608
                  _ExtentX        =   9366
                  _ExtentY        =   3651
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
                  FormatString    =   "No|予定日|予定数|仕入先    |受入|"
                  Rows            =   6
                  Cols            =   6
                  BackColor       =   16777152
                  BackColorBkg    =   14737632
                  FocusRect       =   0
                  HighLight       =   2
                  ScrollBars      =   2
                  RowHeightMin    =   330
                  SelectionMode   =   1
                  BackColorAlternate=   16777152
               End
               Begin Cis3D_v60.CIS3D CIS3D6 
                  Height          =   405
                  Left            =   30
                  Top             =   5280
                  Width           =   5325
                  _ExtentX        =   9393
                  _ExtentY        =   714
                  BackColor       =   16777152
                  Caption         =   "【　入 荷 予 定 情 報  】"
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
                  cAlingnment     =   7
                  Begin Cis3D_v60.CIS3D B1lb_Chuzan 
                     Height          =   345
                     Left            =   30
                     Top             =   30
                     Width           =   1245
                     _ExtentX        =   2196
                     _ExtentY        =   609
                     BackColor       =   16777088
                     Caption         =   "##,###"
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
               Begin Cis3D_v60.CIS3D CIS3D7 
                  Height          =   405
                  Left            =   30
                  Top             =   30
                  Width           =   5325
                  _ExtentX        =   9393
                  _ExtentY        =   714
                  BackColor       =   12648384
                  Caption         =   "【　品　番　情　報  】"
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
                  cAlingnment     =   7
               End
               Begin Cis3D_v60.CIS3D CIS3D27 
                  Height          =   795
                  Left            =   60
                  Top             =   1260
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   1402
                  ForeColor       =   16711680
                  Caption         =   "収容数①"
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
                  Begin Cis3D_v60.CIS3D B1lb_Syuyo1 
                     Height          =   330
                     Left            =   270
                     Top             =   390
                     Width           =   1035
                     _ExtentX        =   1826
                     _ExtentY        =   582
                     BackColor       =   16777152
                     Caption         =   "99,999"
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
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -30
                  End
               End
               Begin Cis3D_v60.CIS3D CIS3D10 
                  Height          =   1590
                  Left            =   1740
                  Top             =   1260
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   2805
                  ForeColor       =   16711680
                  Caption         =   "切替日"
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
                  cPositionY      =   210
                  Begin Cis3D_v60.CIS3D B1lb_CDate 
                     Height          =   330
                     Left            =   240
                     Top             =   750
                     Width           =   1275
                     _ExtentX        =   2249
                     _ExtentY        =   582
                     BackColor       =   16777152
                     Caption         =   "9999/99/99"
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
                     cAlingnment     =   7
                     cBoderStyle     =   1
                  End
                  Begin Cis3D_v60.CIS3D B1lb_CBin 
                     Height          =   330
                     Left            =   1530
                     Top             =   750
                     Width           =   405
                     _ExtentX        =   714
                     _ExtentY        =   582
                     BackColor       =   16777152
                     Caption         =   "99"
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
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -30
                  End
               End
               Begin Cis3D_v60.CIS3D CIS3D16 
                  Height          =   795
                  Left            =   3870
                  Top             =   1260
                  Width           =   1500
                  _ExtentX        =   2646
                  _ExtentY        =   1402
                  ForeColor       =   16711680
                  Caption         =   "収容数②"
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
                  Begin Cis3D_v60.CIS3D B1lb_Syuyo2 
                     Height          =   330
                     Left            =   210
                     Top             =   390
                     Width           =   1035
                     _ExtentX        =   1826
                     _ExtentY        =   582
                     BackColor       =   16777152
                     Caption         =   "99,999"
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
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -30
                  End
               End
               Begin Cis3D_v60.CIS3D CIS3D3 
                  Height          =   795
                  Left            =   60
                  Top             =   2055
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   1402
                  ForeColor       =   16711680
                  Caption         =   "発注ﾛｯﾄ①"
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
                  Begin Cis3D_v60.CIS3D B1lb_Lot1 
                     Height          =   330
                     Left            =   270
                     Top             =   390
                     Width           =   1035
                     _ExtentX        =   1826
                     _ExtentY        =   582
                     BackColor       =   16777152
                     Caption         =   "99,999"
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
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -30
                  End
               End
               Begin Cis3D_v60.CIS3D CIS3D25 
                  Height          =   795
                  Left            =   3870
                  Top             =   2055
                  Width           =   1500
                  _ExtentX        =   2646
                  _ExtentY        =   1402
                  ForeColor       =   16711680
                  Caption         =   "発注ﾛｯﾄ②"
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
                  Begin Cis3D_v60.CIS3D B1lb_Lot2 
                     Height          =   330
                     Left            =   210
                     Top             =   390
                     Width           =   1035
                     _ExtentX        =   1826
                     _ExtentY        =   582
                     BackColor       =   16777152
                     Caption         =   "99,999"
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
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -30
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Week 
                  Height          =   630
                  Left            =   60
                  Top             =   2850
                  Width           =   5325
                  _ExtentX        =   9393
                  _ExtentY        =   1111
                  ForeColor       =   16711680
                  Caption         =   "曜日指定"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   1
                  cPositionX      =   400
                  Begin Cis3D_v60.CIS3D B1lb_Week 
                     Height          =   345
                     Index           =   1
                     Left            =   1965
                     Top             =   135
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
                     BackColor       =   16777152
                     Caption         =   "月"
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
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Week 
                     Height          =   345
                     Index           =   2
                     Left            =   2565
                     Top             =   135
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
                     BackColor       =   16777152
                     Caption         =   "火"
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
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Week 
                     Height          =   345
                     Index           =   3
                     Left            =   3165
                     Top             =   135
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
                     BackColor       =   16777152
                     Caption         =   "水"
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
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Week 
                     Height          =   345
                     Index           =   4
                     Left            =   3765
                     Top             =   135
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
                     BackColor       =   16777152
                     Caption         =   "木"
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
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Week 
                     Height          =   345
                     Index           =   5
                     Left            =   4365
                     Top             =   135
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
                     BackColor       =   16777152
                     Caption         =   "金"
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
                  End
               End
               Begin VB.Label B1lb_ZaiValue 
                  BackStyle       =   0  '透明
                  Caption         =   "XXXXX"
                  Height          =   225
                  Left            =   2790
                  TabIndex        =   25
                  Top             =   930
                  Width           =   570
               End
               Begin VB.Label B1lb_ZaiKB 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '透明
                  Caption         =   "XXXXXX"
                  Height          =   225
                  Left            =   2010
                  TabIndex        =   24
                  Top             =   930
                  Width           =   720
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '透明
                  Caption         =   "在庫発注"
                  Height          =   225
                  Left            =   2220
                  TabIndex        =   23
                  Top             =   570
                  Width           =   960
               End
               Begin VB.Shape Shape6 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   1  '不透明
                  Height          =   390
                  Left            =   1950
                  Shape           =   4  '丸みのある長方形
                  Top             =   855
                  Width           =   1485
               End
               Begin VB.Shape Shape3 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   1  '不透明
                  Height          =   390
                  Left            =   1950
                  Shape           =   4  '丸みのある長方形
                  Top             =   480
                  Width           =   1485
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '透明
                  Caption         =   "予定数"
                  Height          =   225
                  Left            =   4500
                  TabIndex        =   22
                  Top             =   570
                  Width           =   720
               End
               Begin VB.Label B1lb_YSu 
                  Alignment       =   1  '右揃え
                  BackStyle       =   0  '透明
                  Caption         =   "###,##9"
                  Height          =   225
                  Left            =   4440
                  TabIndex        =   21
                  Top             =   930
                  Width           =   840
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  '透明
                  Caption         =   "設定数"
                  Height          =   225
                  Left            =   3570
                  TabIndex        =   20
                  Top             =   570
                  Width           =   735
               End
               Begin VB.Label B1lb_SetSu 
                  Alignment       =   1  '右揃え
                  BackStyle       =   0  '透明
                  Caption         =   "###,##9"
                  Height          =   225
                  Left            =   3480
                  TabIndex        =   19
                  Top             =   930
                  Width           =   840
               End
               Begin VB.Label B1lb_SRyo 
                  Alignment       =   1  '右揃え
                  BackStyle       =   0  '透明
                  Caption         =   "###,##9"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   16
                  Top             =   930
                  Width           =   840
               End
               Begin VB.Shape Shape5 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   1  '不透明
                  Height          =   390
                  Left            =   60
                  Shape           =   4  '丸みのある長方形
                  Top             =   855
                  Width           =   1005
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  '透明
                  Caption         =   "所要量"
                  Height          =   225
                  Left            =   210
                  TabIndex        =   15
                  Top             =   570
                  Width           =   750
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   1  '不透明
                  Height          =   420
                  Left            =   60
                  Shape           =   4  '丸みのある長方形
                  Top             =   480
                  Width           =   1005
               End
               Begin VB.Label Label7 
                  BackStyle       =   0  '透明
                  Caption         =   "現在庫数"
                  Height          =   225
                  Left            =   1170
                  TabIndex        =   18
                  Top             =   570
                  Width           =   795
               End
               Begin VB.Label B1lb_ZaiSu 
                  Alignment       =   1  '右揃え
                  BackStyle       =   0  '透明
                  Caption         =   "##,##9"
                  Height          =   225
                  Left            =   1170
                  TabIndex        =   17
                  Top             =   930
                  Width           =   750
               End
               Begin VB.Shape Shape4 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   1  '不透明
                  Height          =   390
                  Left            =   1050
                  Shape           =   4  '丸みのある長方形
                  Top             =   480
                  Width           =   915
               End
               Begin VB.Shape Shape2 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   1  '不透明
                  Height          =   390
                  Left            =   1050
                  Shape           =   4  '丸みのある長方形
                  Top             =   855
                  Width           =   915
               End
               Begin VB.Shape Shape12 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   1  '不透明
                  Height          =   390
                  Left            =   4350
                  Shape           =   4  '丸みのある長方形
                  Top             =   480
                  Width           =   1035
               End
               Begin VB.Shape Shape11 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   1  '不透明
                  Height          =   390
                  Left            =   4350
                  Shape           =   4  '丸みのある長方形
                  Top             =   855
                  Width           =   1035
               End
               Begin VB.Shape Shape10 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   1  '不透明
                  Height          =   390
                  Left            =   3420
                  Shape           =   4  '丸みのある長方形
                  Top             =   480
                  Width           =   945
               End
               Begin VB.Shape Shape9 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   1  '不透明
                  Height          =   390
                  Left            =   3420
                  Shape           =   4  '丸みのある長方形
                  Top             =   855
                  Width           =   945
               End
            End
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
      Height          =   975
      Left            =   135
      ScaleHeight     =   915
      ScaleWidth      =   9675
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1230
      Width           =   9735
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   810
         Left            =   3405
         Top             =   60
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1429
         ForeColor       =   16711680
         Caption         =   "仕 入 先"
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
            Left            =   75
            TabIndex        =   2
            Top             =   345
            Width           =   990
            _ExtentX        =   1746
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
            Height          =   360
            Left            =   1065
            Top             =   345
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XX"
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
         Height          =   810
         Index           =   1
         Left            =   8280
         Top             =   60
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1429
         ForeColor       =   16711680
         Caption         =   "発注残"
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
         Begin VB.OptionButton OpH1_HZan 
            BackColor       =   &H00C0FFC0&
            Caption         =   "含む"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   75
            TabIndex        =   4
            Top             =   255
            Width           =   1050
         End
         Begin VB.OptionButton OpH1_HZan 
            BackColor       =   &H00C0C0C0&
            Caption         =   "含まない"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   75
            TabIndex        =   5
            Top             =   525
            Width           =   1170
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   810
         Index           =   1
         Left            =   6300
         Top             =   60
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   1429
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D H1lb_TantXNm 
            Height          =   360
            Left            =   585
            Top             =   330
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXXX"
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
         Begin CisText_V60.CisText iH1_TantX 
            Height          =   360
            Left            =   90
            TabIndex        =   3
            Top             =   330
            Width           =   495
            _ExtentX        =   873
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
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   810
         Left            =   90
         Top             =   60
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1429
         ForeColor       =   16711680
         Caption         =   "品   番"
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
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   60
            TabIndex        =   1
            Top             =   330
            Width           =   3240
            _ExtentX        =   5715
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
      Left            =   11670
      ScaleHeight     =   660
      ScaleWidth      =   3480
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   420
      Width           =   3540
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   915
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
         Left            =   1755
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
         Left            =   2595
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
      IMEMode         =   3  'ｵﾌ固定
      Left            =   960
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   -1320
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   900
      Top             =   -1380
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 発注指示 (発注点)  】"
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
         Left            =   14400
         Top             =   75
         Width           =   915
         _ExtentX        =   1614
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
      Begin Cis3D_v60.CIS3D UC_3D32 
         Height          =   255
         Left            =   8565
         Top             =   60
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
   End
   Begin VB.PictureBox H1_Area2 
      Height          =   975
      Left            =   135
      ScaleHeight     =   915
      ScaleWidth      =   2115
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   2175
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   810
         Index           =   0
         Left            =   90
         Top             =   60
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   1429
         ForeColor       =   16711680
         Caption         =   "発注担当者"
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
         Begin Cis3D_v60.CIS3D H1lb_TantNm 
            Height          =   360
            Left            =   585
            Top             =   330
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXXX"
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
         Begin CisText_V60.CisText iH1_Tant 
            Height          =   360
            Left            =   90
            TabIndex        =   0
            Top             =   330
            Width           =   495
            _ExtentX        =   873
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
   End
End
Attribute VB_Name = "CXH0080"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   発注指示(発注点)
'**       フォームID    :   CXH0080
'**       処理概要      :
'**
'**       作  成  日    :   2004/11/22  By CIS
'**       変  更  日    :   2005/03/30  By CIS  担当者の検索結果を画面上に反映
'**       変  更  日    :   2005/11/17  By CIS オプション追加（稼働日カレンダー）
'**                                             発注日算出=自社カレンダー、納入日算出=仕入先カレンダー
'**       変  更  日    :   2007/07/13  By CIS 仕入先･受入カレンダー対応
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid               As New CisVsGrid3
    Dim CisVsGridH              As New CisVsGrid3
    Dim CisVsGridY              As New CisVsGrid3
    
    Dim mGotColor               As Long
    Dim mLostColor              As Long
    
    Dim mWhere                  As String
    
    Dim ActObj                  As Object
    Dim ActRow                  As Integer
    
    Dim w_RowMax                As Long
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
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
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
    
    mGotColor = OpH1_HZan(0).BackColor
    mLostColor = OpH1_HZan(1).BackColor
    
    OpH1_HZan(0).Value = True
    If InStr(1, UCase(Command), "/N") <> 0 Or InStr(1, UCase(Command), "/NO") <> 0 Then
       OpH1_HZan(1).Value = True
       OpH1_HZan(0).BackColor = mLostColor
       OpH1_HZan(1).BackColor = mGotColor
    End If
    
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 380
        .PatanMax = 1
        .SelectCol = 5
        .EditCheckCol = 1
        'Set .SelectDispObj = B1lb_SelSu
        .InitGet = False
        .Init
    End With
    Call GridInit
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGridY
        Set .GridObj = VsGridY
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .InitGet = False
        .Init
    End With
    Call GridInitY
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGridH
        Set .GridObj = VsGridH
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .InitGet = False
        .Init
    End With
    Call GridInitH
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
    
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'システムオプション獲得
    Call GetSYSOption(33)
    
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
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
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
'+    取  消  キ ー(Escape)   +
'+----------------------------+
Private Sub PB_Can_Click()
    Call CisFrm.MousePT(11)
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   
    If ProcHB = "H1" Then
       Call HeadBodyClear("H")
       GoTo PB_CAN_Ed
    End If
     
    If ProcHB = "B1" Then
       ProcHB = "H1"
       Call CisVsGrid.EditEnable
       Call DispChange(ProcHB)
       Call HeadBodyClear("B")
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
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_SSTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Tant" Then Call Look_Tanto: Exit Sub
'    If PB_Look.Tag = "iH1_TantX" Then Call Look_Tanto: Exit Sub        '2005.03.30 Upd
    If PB_Look.Tag = "iH1_TantX" Then Call Look_TantoX: Exit Sub
    If PB_Look.Tag = "iB1_NYmd" Then iB1_NYmd.ShowCalender: iB1_NYmd.SetFocus: Exit Sub

End Sub
'+----------------------------+
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0
    RV_Top = 0
    RVI_Hinmoku = 1
    RVI_SK = 9
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iH1_Hinbn = RV_Hinbn
    End If
End Sub
'+-----------------------------+
'+       取引先検索             +
'+-----------------------------+
Private Sub Look_SSTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+----------------------------+
'+     発注担当者検索         +
'+----------------------------+
Private Sub Look_Tanto()

    RV_Left = 0
    RV_Top = 0
    RVI_Kubun = "担当管理"

    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        iH1_Tant = RV_Code
        H1lb_TantNm = RV_Kubun
    End If

End Sub
'+----------------------------+
'+       担当者検索           +
'+----------------------------+
Private Sub Look_TantoX()

    RV_Left = 0
    RV_Top = 0
    RVI_Kubun = "担当管理"

    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        iH1_TantX = RV_Code
        H1lb_TantXNm = RV_Kubun
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
'+  iH1_Tant
'+-----------------------------+
Private Sub iH1_Tant_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Tant_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_TantNm = ""
    If RTrim(iH1_Tant) <> "" Then
        If MeisyoGet("担当管理", iH1_Tant) Then
            H1lb_TantNm = SYM_Meisyo
        End If
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
    If TorNmGet(iH1_Torcd, 8) Then
        H1lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iH1_TantX
'+-----------------------------+
Private Sub iH1_TantX_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_TantX_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_TantXNm = ""
    If RTrim(iH1_TantX) <> "" Then
        If MeisyoGet("担当管理", iH1_TantX) Then
            H1lb_TantXNm = SYM_Meisyo
        End If
    End If
End Sub
'+----------------------------------+
'+      OpH1_HZan
'+----------------------------------+
Private Sub OpH1_HZan_GotFocus(Index As Integer)
    For Each gObj In OpH1_HZan
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
'+----------------------+
'+  iB1_HSuryo
'+----------------------+
Private Sub iB1_HSuryo_GotFocus()
    Set ActObj = iB1_HSuryo
End Sub
Private Sub iB1_HSuryo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
    Call GridToDisp
End Sub
Private Sub iB1_HSuryo_LostFocus()
    If CisFun.Val2(iB1_HSuryo) <> 0 Then
        vsGrid1.TextMatrix(CisFun.Val2(iB1_HSuryo.Tag), 2) = iB1_HSuryo
    Else
        vsGrid1.TextMatrix(CisFun.Val2(iB1_HSuryo.Tag), 2) = ""
    End If
End Sub
'+----------------------+
'+  iB1_NYmd
'+----------------------+
Private Sub iB1_NYmd_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
    Call GridToDisp
End Sub
Private Sub iB1_NYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_NYmd
End Sub
Private Sub iB1_NYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    If Trim(iB1_NYmd) <> "" Then
        vsGrid1.TextMatrix(CisFun.Val2(iB1_HSuryo.Tag), 3) = iB1_NYmd.cYear & "/" & iB1_NYmd.cMonth & "/" & iB1_NYmd.cDay
    Else
        vsGrid1.TextMatrix(CisFun.Val2(iB1_HSuryo.Tag), 3) = ""
    End If
End Sub
'+--------------------------------+
'+  グリッドから表示
'+--------------------------------+
Private Sub GridToDisp()
    Dim wHinbn      As String
    Dim wTorcd      As String
    Dim wUkeir      As String
    
    
    ' 品番情報表示
    With vsGrid1
        lb_Kensu.Caption = Format(.TextMatrix(.Row, CisVsGrid.FixedGet(0)), "###,##0") & " / " & _
                           Format(w_RowMax, "###,##0")
         
         ' 所要量
         B1lb_SRyo = Format(.TextMatrix(.Row, CisVsGrid.FixedGet(7)), "##,###")
         ' 現在庫数
         B1lb_ZaiSu = Format(.TextMatrix(.Row, CisVsGrid.FixedGet(8)), "##,###")
         ' 在庫発注
         B1lb_ZaiKB = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(9)))
         ' 在庫発注値
         B1lb_ZaiValue = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(19)))
         ' 設定数
         B1lb_SetSu = Format(.TextMatrix(.Row, CisVsGrid.FixedGet(10)), "##,###")
         ' 発注予定数
         B1lb_YSu = Format(.TextMatrix(.Row, CisVsGrid.FixedGet(11)), "##,###")
         ' 収容数/発注ロット 1
         B1lb_Syuyo1 = Format(.TextMatrix(.Row, CisVsGrid.FixedGet(12)), "##,###")
         B1lb_Lot1 = Format(.TextMatrix(.Row, CisVsGrid.FixedGet(13)), "##,###")
         ' 切替日・便
         B1lb_CDate = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(14)))
         B1lb_CBin = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(15)))
         ' 収容数/発注ロット 2
         B1lb_Syuyo2 = Format(.TextMatrix(.Row, CisVsGrid.FixedGet(16)), "##,###")
         B1lb_Lot2 = Format(.TextMatrix(.Row, CisVsGrid.FixedGet(17)), "##,###")
         ' 指定
         Back_Week = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(18)))
         B1lb_Week(1) = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(20)))
         B1lb_Week(2) = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(21)))
         B1lb_Week(3) = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(22)))
         B1lb_Week(4) = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(23)))
         B1lb_Week(5) = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(24)))
        
         wHinbn = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(25)))
         wTorcd = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(26)))
         wUkeir = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(5)))
    End With
    ' 発注先情報
    Call HInfoDisp(wHinbn, wTorcd, wUkeir)
    ' 入荷予定情報
    Call NInfoDisp(wHinbn, wTorcd, wUkeir)
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
    Dim B2Mode      As Boolean
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    B2Mode = False
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
        Case "B2":      B2Mode = True
    End Select
    
    H1_Area2.Enabled = H1Mode
    H1_Area2.BackColor = H1Color
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    PB_Look.Visible = H1Mode
    
    Call GridTextEnabled(B1Mode)
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
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        lb_Kensu = ""
        CisVsGrid.Clear
        CisVsGridH.Clear
        CisVsGridY.Clear
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
       ProcHB = "B1"
       Call DispChange(ProcHB)
       GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
       If Not Body1Chk Then GoTo ReturnPress_Ed
    ' << 発注処理 >>
        If Not HCProc Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
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
    
'   >> 担当者
    H1lb_TantNm = ""
    If Trim(iH1_Tant) <> "" Then
        If MeisyoGet("担当管理", iH1_Tant) Then
            H1lb_TantNm = SYM_Meisyo
        Else
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    発注担当者　未登録！（名称マスタ）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Tant.SetFocus
            Exit Function
        End If
    End If
    
    Message.Visible = True
    MessageDisp.Caption = "【抽出中】しばらくお待ち下さい"
    DoEvents

' 発注点発注ワーク 抽出
    If Not HCExtraction Then
       Message.Visible = False
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データがありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
       iH1_Tant.SetFocus
       Exit Function
    End If
    
' 発注点発注ワーク SELECT文 ｾｯﾄ
    gSL_Select = "Select HH.*," & vbCrLf
    gSL_Select = gSL_Select & "HN.在庫管理区分,HN.在庫係数,HN.在庫日数,HN.在庫発注点," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HN.設定数,0) 設定数,HN.表示品番,HN.担当," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(SY.値名称,'') 在庫管理名," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.背番号,'') 背番号," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.手配区分,'') 手配区分," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.SK区分,0) SK区分," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.収容数1,0) 収容数1," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.ロット1,0) 発注ロット1," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.変更日,'') 切替日," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.変更便,'') 切替便," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.収容数2,0) 収容数2," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.ロット2,0) 発注ロット2," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.Week1,0) Week1," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.Week2,0) Week2," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.Week3,0) Week3," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.Week4,0) Week4," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.Week5,0) Week5," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.Day1,0) Day1," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.Day2,0) Day2," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.Day3,0) Day3," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.Day4,0) Day4," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(HT.Day5,0) Day5," & vbCrLf
    gSL_Select = gSL_Select & "    ISNULL(TR.略称,'') 仕入先名" & vbCrLf

    gSL_Select = gSL_Select & "    From 発注点発注ワーク HH " & vbCrLf
    gSL_Select = gSL_Select & "    Left Outer Join 品番マスタ HN " & vbCrLf
    gSL_Select = gSL_Select & "     ON  HN.品番 = HH.品番 " & vbCrLf
    gSL_Select = gSL_Select & "    Left Outer Join 名称マスタ SY " & vbCrLf
    gSL_Select = gSL_Select & "     ON  SY.区分名称 = '在庫区分' " & vbCrLf
    gSL_Select = gSL_Select & "     And SY.区分タイプ = '' " & vbCrLf
    gSL_Select = gSL_Select & "     And SY.値 = HN.在庫管理区分 " & vbCrLf
    gSL_Select = gSL_Select & "    Left Outer Join 品番取引先マスタ HT " & vbCrLf
    gSL_Select = gSL_Select & "     ON  HT.品番 = HN.品番 " & vbCrLf
    gSL_Select = gSL_Select & "     AND HT.品目 = 1 " & vbCrLf
    gSL_Select = gSL_Select & "     And HT.取引先 = HH.仕入先 " & vbCrLf
    gSL_Select = gSL_Select & "     AND HT.受入  = HH.受入 " & vbCrLf
    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TR " & vbCrLf
    gSL_Select = gSL_Select & "     ON  TR.取引先CD = HH.仕入先 " & vbCrLf
    gSL_Select = gSL_Select & "     AND TR.取引先区分 = 1 " & vbCrLf

    ' 抽出条件
    mWhere = ""
    With CisFun
        ' 発注予定数
        .Where_CreateKB = NewDocument
        .Where_Create Numeric, "HH.発注予定数", "0", 大きい, Create
        ' 品番
'        .Where_Create AlphaMe, "HH.品番", iH1_Hinbn
        ' 仕入先
        .Where_Create AlphaMe, "HH.仕入先", iH1_Torcd
        ' 担当者
'        .Where_Create AlphaMe, "HN.担当", iH1_TantX
        gSL_Select = gSL_Select & .Where_Phrase
    End With

'< 2004.11.16 ADD START >
    gSL_Select = gSL_Select & " AND   HH.処理端末 = Host_Name()"
'< 2004.11.16 ADD  END  >
    
    gSL_Select = gSL_Select & " Order By HH.品番,HH.仕入先,HH.受入"
'
    Message.Visible = True
    MessageDisp.Caption = "【表示中】しばらくお待ち下さい"
    DoEvents
    
    If Not HHWRead(gSL_Select) Then
       Message.Visible = False
       Call HHWClose
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データがありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
       iH1_Hinbn.SetFocus
       Exit Function
    End If
    
    w_RowMax = CisDB.RecordCount
    
    Call GridSet
    
' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
    vsGrid1.Row = 1
    Call VsGridToText(vsGrid1.Row)
    
    Message.Visible = False
    
    Call GridToDisp
    
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Dim ChkFlg      As Boolean
    Dim wNYMD       As String
    Dim wHSu        As Long
    Dim wErGyo      As Long
    Dim wErMsg      As String
    
    Body1Chk = False
    
    Call TextToVsGrid(Val(iB1_HSuryo.Tag))
    
    With vsGrid1
        ChkFlg = False
        wErGyo = 0
        wErMsg = ""
        
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(1))) = "" Then Exit For
            
            wHSu = CisFun.Val2(.TextMatrix(gLong, CisVsGrid.FixedGet(2)))
            wNYMD = Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(3)))
            If wNYMD <> "" Then
               If Not DateChk(wNYMD) Then
                  wErGyo = gLong
                  wErMsg = "正しい日付にて入力して下さい"
                  Set gObj = iB1_NYmd
                  GoTo Body1Chk_Er
               End If
               If RTrim(wNYMD) < Format(Now(), "yyyymmdd") Then
                  With CisFun
                       .MB_Lines = 5
                       .MB_MSG(2) = "     " & StrConv(gLong, vbWide) & "行目 納入日が過去の日付です。     "
                       .MB_MSG(4) = "        よろしいですか？          "
                       .MB_Title = ""
                       .MB_Button = OK_CAN
                       If Not .MBOX Then
                          wErGyo = gLong
                          Set gObj = iB1_NYmd
                          GoTo Body1Chk_E2
                       End If
                   End With
                End If
                If gSYSOp_CLKb = 1 Then
                    RV_Torcd = Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(26)))
                    RV_Ukeir = Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(5)))
                    gInt = Kdou_Check(wNYMD, 1, RV_Torcd, RV_Ukeir)           '仕入先カレンダー
                    If gInt = 2 Then
                        gInt = Kdou_Check(wNYMD)
                    End If
                Else
                    gInt = Kdou_Check(wNYMD)
                End If
                If gInt = 1 Then
                '休日
                   With CisFun
                       .MB_Lines = 5
                       .MB_MSG(2) = "     " & StrConv(gLong, vbWide) & "行目 納入日は非稼働日です。     "
                       .MB_MSG(4) = "        よろしいですか？          "
                       .MB_Title = ""
                       .MB_Button = OK_CAN
                       If Not .MBOX Then
                          wErGyo = gLong
                          Set gObj = iB1_NYmd
                          GoTo Body1Chk_E2
                       End If
                    End With
               End If
               If wHSu = 0 Then
                  wErGyo = gLong
                  wErMsg = "発注数を入力して下さい"
                  Set gObj = iB1_HSuryo
                  GoTo Body1Chk_Er
               End If
               ChkFlg = True
            End If
        Next gLong
    End With
        
    
    If Not ChkFlg Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    納入日を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        wErGyo = 1
        Set gObj = iB1_NYmd
        GoTo Body1Chk_E2
    End If

    Body1Chk = True
    Exit Function
Body1Chk_Er:
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "    " & StrConv(wErGyo, vbWide) & "行目 " & wErMsg & "            "
        .MB_Title = ""
        .MB_Button = Error
        .MBOX
    End With
Body1Chk_E2:
    vsGrid1.Row = wErGyo
    Call VsGridToText(vsGrid1.Row)
    Call GridTextEnabled(True)
    gObj.SetFocus
    
End Function
Private Function DateChk(ChkYmd As String, Optional YMD As Date)
    Dim yy As String
    Dim mm As String
    Dim DD As String
    
    DateChk = False: YMD = 0
    
    If Trim(ChkYmd) = "" Then
       DateChk = True
       Exit Function
    End If
    
    ChkYmd = Replace(ChkYmd, "/", "")
    
    yy = Mid(ChkYmd, 1, 4)
    mm = Mid(ChkYmd, 5, 2)
    DD = Mid(ChkYmd, 7, 2)
    
    On Error GoTo YMD_ERR:

    If yy < 95 Then
        yy = 2000 + yy
    Else
        If yy < 100 Then
            yy = 1900 + yy
        End If
    End If
    YMD = DateValue(yy & "," & mm & "," & DD)
    DateChk = True
    
    On Error GoTo 0
    
    Exit Function
YMD_ERR:
    On Error GoTo 0
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'//////////////////////////////////////
'//
'//     発注点データ　抽出
'//
'//////////////////////////////////////
Private Function HCExtraction() As Boolean
    HCExtraction = False

    On Error GoTo HCExtraction_ER
    gStr = ""
    With CisDB
        .SQL = "発注点データ抽出"
        .StoadoCount = 5
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = Input用
        .ParaNo = 4: .ParaIO = OutPut用
        .ParaNo = 5: .ParaIO = OutPut用

'<品番>
        .ParaNo = 1: .ParaValue = RTrim(iH1_Hinbn)
'<担当者>
        .ParaNo = 2: .ParaValue = RTrim(iH1_TantX)
        
        If OpH1_HZan(0).Value Then
           .ParaNo = 3: .ParaValue = 0
        Else
           .ParaNo = 3: .ParaValue = 1
        End If

        If Not .DBStored Then GoTo HCExtraction_ER

        ' ストアド処理エラー
        .ParaNo = 4: gCnt1 = CisFun.Val2(.ParaValue)
        If gCnt1 <> 0 Then GoTo HCExtraction_ER

    End With

    HCExtraction = True
    On Error GoTo 0
    Exit Function
HCExtraction_ER:
    Select Case gCnt1
        Case 1: gStr = "発注点発注ワーク 削除エラー"
        Case 10: gStr = "品番マスタ 更新エラー"
        Case 90: gStr = "発注点発注ワーク 作成エラー"
        Case Else
             gStr = "その他ｴﾗｰ:" & gStr
    End Select
    gStr = StrConv(Format(gCnt1, "00"), vbWide) & ":" & gStr
    Call CisFun.ErrorBox(gStr)

End Function
'+----------------------+
'+    発注処理 確認     +
'+----------------------+
Private Function HCProc() As Boolean
    HCProc = False
    Dim AddFlg  As Boolean

    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     発注処理を行います      "
        .MB_Title = "発注処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo HCProc_Er

    Call CisVsGrid.EditEnable
    DoEvents

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
    With vsGrid1
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(1))) = "" Then Exit For

            If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(3))) <> "" Then
                If Not HCPut(gLong) Then Exit Function
            End If
        Next gLong
    End With

    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
'
    Call EndProc
    
    HCProc = True
    
    
    Exit Function
HCProc_Er:
    Call CisDB.DBTran(TransRollback)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 破棄
    CisFun.ErrorBox
End Function
'+------------------------------------------+
'+      発注画面入力ワーク出力
'+------------------------------------------+
Private Function HCPut(wRow As Long) As Boolean
    HCPut = False
    
    Call ItemsClearHGW
    With HGW
        .発注区分 = "5"
        .仕入先 = RTrim(vsGrid1.TextMatrix(wRow, CisVsGrid.FixedGet(26)))
        .受入 = RTrim(vsGrid1.TextMatrix(wRow, CisVsGrid.FixedGet(5)))
        gStr = RTrim(vsGrid1.TextMatrix(wRow, CisVsGrid.FixedGet(3)))
        gStr = Replace(gStr, "/", "")
        .納入日 = gStr
        .品番 = RTrim(vsGrid1.TextMatrix(wRow, CisVsGrid.FixedGet(25)))
        .背番号 = RTrim(vsGrid1.TextMatrix(wRow, CisVsGrid.FixedGet(6)))
        .手配区分 = RTrim(vsGrid1.TextMatrix(wRow, CisVsGrid.FixedGet(27)))
        .SK区分 = CisFun.Val2(vsGrid1.TextMatrix(wRow, CisVsGrid.FixedGet(28)))
'        .サイクル = CisFun.Val2(vsGrid.TextMatrix(gIndex, 12))
'        .契約先 = CisFun.Val2(vsGrid.TextMatrix(gIndex, 13))
        .収容数 = CisFun.Val2(vsGrid1.TextMatrix(wRow, CisVsGrid.FixedGet(12)))
        gStr = RTrim(vsGrid1.TextMatrix(wRow, CisVsGrid.FixedGet(14)))
        gStr = Replace(gStr, "/", "")
        If gStr <> "" Then
            If gStr <= .納入日 Then
               .収容数 = CisFun.Val2(vsGrid1.TextMatrix(wRow, CisVsGrid.FixedGet(16)))
            End If
        End If
        .発注数 = CisFun.Val2(vsGrid1.TextMatrix(wRow, CisVsGrid.FixedGet(2)))
        If .収容数 <> 0 Then
            gCnt1 = .発注数 \ .収容数
            gCnt2 = .発注数 Mod .収容数
            If gCnt2 <> 0 Then
                gCnt1 = gCnt1 + 1
            End If
            .枚数 = gCnt1
            .端数 = gCnt2
        End If
        .作成種別 = "Y"
        .処理区分 = 0
'        .担当者 = RTrim$(iH1_Tant)
        .作成者 = gTanto
    End With
 ' ( 追加処理 )
    If Not HGWInsert Then Exit Function
    
    HCPut = True

End Function
'==========================================================='
'   終了処理（入力データがあれば後方処理を行う）
'==========================================================='
Private Function EndProc() As Boolean
    Dim PrtCnt3     As Long
    Dim PrtCnt4     As Long
'入力件数確認
    EndProc = True
    gSL_Select = "SELECT COUNT(*) 件数 FROM 発注画面入力ワーク"
    gSL_Select = gSL_Select & " WHERE 処理端末 = HOST_NAME()"
    If Not HGWRead(gSL_Select, 1) Then Exit Function
    If HGW.件数 = 0 Then Exit Function

    EndProc = False
'発注打切更新
    With CisDB
        .SQL = "発注画面入力後方処理"
        .StoadoCount = 5
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = OutPut用            '入力件数
        .ParaNo = 2: .ParaIO = OutPut用            '発注打切テーブル更新件数
        .ParaNo = 3: .ParaIO = OutPut用            '発注打切警告リスト件数
        .ParaNo = 4: .ParaIO = OutPut用            '発注打切警告リスト件数
        .ParaNo = 5: .ParaIO = OutPut用            'エラーコード
'
        If Not .DBStored Then
            MsgBox "ストアドエラー "
            GoTo EndProc_Err
        Else
            .ParaNo = 5
            If .ParaValue <> 0 Then GoTo EndProc_Err
            .ParaNo = 3
            PrtCnt3 = .ParaValue
            .ParaNo = 4
            PrtCnt4 = .ParaValue
        End If
    End With
'発注打切警告リスト印刷
    If PrtCnt3 <> 0 Then
        CHH0930_CallType = ""
        If Not CKH0930LPrint Then Exit Function
        DoEvents
    End If
'発注打切リスト印刷
    If PrtCnt4 <> 0 Then
        CHH0940_CallType = ""
        If Not CKH0940LPrint Then Exit Function
        DoEvents
    End If
' 発注処理
    CisDB.PassLink
    With CisFun
        .ShellApl = gIniExe & "PXH0110.exe"
        .ShellAplPara = "G"
        .ShellTask Freeze
    End With

    EndProc = True

    Exit Function
EndProc_Err:
    Select Case CisDB.ParaValue
        Case 1: gStr = "    【　かんばん読取エラーテーブル　削除エラー！】"
        Case 2: gStr = "    【　帳票出力エラーテーブル　削除エラー！】"
        Case 3: gStr = "    【　発注打切テーブル(完了)　更新エラー！】"
        Case 4: gStr = "    【　発注画面入力ワーク(端数)　更新エラー！】"
        Case 5: gStr = "    【　発注打切テーブル(未完)　更新エラー！】"
        Case 6: gStr = "    【　帳票出力テーブル(警告)　出力エラー！】"
        Case 7: gStr = "    【　帳票出力テーブル(打切)　出力エラー！】"
        Case 8: gStr = "    【　かんばん読取エラーテーブル　出力エラー！】"
        Case 9: gStr = "    【　発注画面入力ワーク(処理)　更新エラー！】"
        Case 60: gStr = "    【　エラー管理番号　獲得エラー！】"
    End Select
    MsgBox gStr & "  Code = " & CisDB.ParaValue & "     "

End Function

'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    iB1_HSuryo.Visible = EnabledFlg
    iB1_NYmd.Visible = EnabledFlg
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(iB1_HSuryo.Tag))
    gLong = vsGrid1.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gLong - 1 < 1 Then GoTo Grid_RowChange_End
        gLong = gLong - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gLong + 1 > vsGrid1.Rows - 1 Then
           If gLong + 1 < vsGrid1.Rows Then
              If RTrim(vsGrid1.TextMatrix(gLong + 1, 1)) <> "" Then
                 vsGrid1.Rows = gLong + 1 + 1
                 vsGrid1.TextMatrix(gLong + 1, 0) = gLong + 1
              Else
                  GoTo Grid_RowChange_End
              End If
           End If
        End If
        If gLong + 1 < vsGrid1.Rows Then
           If RTrim(vsGrid1.TextMatrix(gLong + 1, 1)) <> "" Then
              gLong = gLong + 1
           End If
        End If
    End If
    'PageUp
    If KeyCode = vbKeyPageUp Then
        If gLong - CisVsGrid.DispMax < 1 Then
            gLong = 1
        Else
            gLong = gLong - CisVsGrid.DispMax
        End If
    End If
    'PageDown
    If KeyCode = vbKeyPageDown Then
        If gLong + CisVsGrid.DispMax > vsGrid1.Rows - 1 Then
            gLong = vsGrid1.Rows - 1
        Else
            gLong = gLong + CisVsGrid.DispMax
        End If
    End If
    vsGrid1.Row = gLong
'グリッドからテキストに表示
    Call VsGridToText(vsGrid1.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    If tRow = 0 Then tRow = 1
    With vsGrid1
        If CisFun.Val2(iB1_HSuryo) <> 0 Then
           .TextMatrix(tRow, 2) = Format(CisFun.Val2(iB1_HSuryo), "#,###")
        Else
           .TextMatrix(tRow, 2) = ""
        End If
        If Trim(iB1_NYmd) <> "" Then
           .TextMatrix(tRow, 3) = iB1_NYmd.cYear & "/" & iB1_NYmd.cMonth & "/" & iB1_NYmd.cDay
        Else
           .TextMatrix(tRow, 3) = ""
        End If
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    If tRow = 0 Then tRow = 1
    With vsGrid1
        iB1_HSuryo.Tag = tRow
        '発注数
        .Col = 2
        iB1_HSuryo.Move .CellLeft, .CellTop, iB1_HSuryo.Width, .CellHeight
        iB1_HSuryo = CisFun.Val2(.TextMatrix(tRow, 2))
        '納入日
        .Col = 3
        iB1_NYmd.Move .CellLeft, .CellTop, iB1_NYmd.Width, .CellHeight
        iB1_NYmd = Replace(.TextMatrix(tRow, 3), "/", "")
    End With
End Sub
Private Sub vsGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    Call TextToVsGrid(Val(iB1_HSuryo.Tag))
    Call VsGridToText(vsGrid1.Row)
End Sub
Private Sub vsGrid1_Click()
    Call TextToVsGrid(Val(iB1_HSuryo.Tag))
    If vsGrid1.TextMatrix(vsGrid1.Row, 0) = "" Then
       vsGrid1.Row = Val(iB1_HSuryo.Tag)
    End If
    Call GridTextEnabled(True)
    Call VsGridToText(vsGrid1.Row)
    Call GridToDisp
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
'Private Sub vsGrid1_Scroll()
'End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, なし, 4, 0, "№")
        Call .FixedSet(中中, 左中, あり, 25, 1, "品番")
        Call .FixedSet(中中, 右中, なし, 10, 2, "発注数")
        Call .FixedSet(中中, 中中, なし, 11, 3, "納入日")
        Call .FixedSet(中中, 左中, なし, 18, 4, "仕入先")
        Call .FixedSet(中中, 左中, なし, 2, 5, "受")
        Call .FixedSet(中中, 左中, あり, 0, 6, "背番号")
        Call .FixedSet(中中, 左中, あり, 0, 7, "所要量")
        Call .FixedSet(中中, 左中, あり, 0, 8, "現在庫数")
        Call .FixedSet(中中, 左中, あり, 0, 9, "在庫発注")
        Call .FixedSet(中中, 左中, あり, 0, 10, "在庫設定値")
        Call .FixedSet(中中, 左中, あり, 0, 11, "発注予定数")
        Call .FixedSet(中中, 左中, あり, 0, 12, "収容数1")
        Call .FixedSet(中中, 左中, あり, 0, 13, "発注ロット1")
        Call .FixedSet(中中, 左中, あり, 0, 14, "切替日")
        Call .FixedSet(中中, 左中, あり, 0, 15, "切替便")
        Call .FixedSet(中中, 左中, あり, 0, 16, "収容数2")
        Call .FixedSet(中中, 左中, あり, 0, 17, "発注ﾛｯﾄ2")
        Call .FixedSet(中中, 左中, あり, 0, 18, "指定日")
        Call .FixedSet(中中, 左中, あり, 0, 19, "在庫発注値")
        Call .FixedSet(中中, 左中, あり, 0, 20, "指定1")
        Call .FixedSet(中中, 左中, あり, 0, 21, "指定2")
        Call .FixedSet(中中, 左中, あり, 0, 22, "指定3")
        Call .FixedSet(中中, 左中, あり, 0, 23, "指定4")
        Call .FixedSet(中中, 左中, あり, 0, 24, "指定5")
        Call .FixedSet(中中, 左中, あり, 0, 25, "検索品番")
        Call .FixedSet(中中, 左中, あり, 0, 26, "検索仕入先")
        Call .FixedSet(中中, 左中, あり, 0, 27, "手配区分")
        Call .FixedSet(中中, 左中, あり, 0, 28, "SK区分")
        Call .FixedSet(中中, 左中, あり, 0, 29, "予備1")
        Call .FixedSet(中中, 左中, あり, 0, 30, "予備2")
        Call .FixedSet(中中, 左中, あり, 0, 31, "予備3")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iB1_HSuryo, 2, , "#,###")
'        Call .EditSet(iB1_NYmd, 3, , , 日付編集あり)
        .InitDisp
    End With
End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ(予定情報)
'+----------------------------------------------------------+
Private Sub GridInitY()
    With CisVsGridY
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "No")
        Call .FixedSet(中中, 中中, あり, 10, 1, "予定日")
        Call .FixedSet(中中, 右中, あり, 8, 2, "予定数")
        Call .FixedSet(中中, 左中, あり, 16, 3, "仕入先")
        Call .FixedSet(中中, 中中, あり, 2, 4, "受")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 5, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ(品番取引先情報)
'+----------------------------------------------------------+
Private Sub GridInitH()
    With CisVsGridH
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "No")
        Call .FixedSet(中中, 左中, あり, 29, 1, "仕入先")
        Call .FixedSet(中中, 中中, あり, 2, 2, "受")
        Call .FixedSet(中中, 右中, あり, 6, 3, "比率")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 4, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
    
    CisVsGrid.Clear

' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HHW_RDSTS
        With vsGrid1
            ID = ID + 1
            
            ' 検索品番
            .TextMatrix(ID, GridCol(25)) = HHW.品番
            ' 検索仕入先
            .TextMatrix(ID, GridCol(26)) = HHW.仕入先
            ' 手配区分
            .TextMatrix(ID, GridCol(27)) = HHW.手配区分
            ' SK区分
            .TextMatrix(ID, GridCol(28)) = HHW.SK区分
            ' 品番
            .TextMatrix(ID, GridCol(1)) = HHW.表示品番
            ' 発注予定数
            .TextMatrix(ID, GridCol(2)) = HHW.発注予定数
            ' 仕入先
            gStr = HHW.仕入先
            If Trim(HHW.仕入先名) <> "" Then
                gStr = gStr & ":" & HHW.仕入先名
            End If
            .TextMatrix(ID, GridCol(4)) = gStr
            ' 受入
            .TextMatrix(ID, GridCol(5)) = HHW.受入
        
            ' 背番号
            .TextMatrix(ID, GridCol(6)) = HHW.背番号
            ' 所要量
            .TextMatrix(ID, GridCol(7)) = HHW.所要量
            ' 現在庫数
            .TextMatrix(ID, GridCol(8)) = HHW.現在庫数
            ' 在庫発注
            .TextMatrix(ID, GridCol(9)) = HHW.在庫管理名
            ' 在庫設定値
            .TextMatrix(ID, GridCol(10)) = HHW.設定数
            ' 在庫発注値
            gStr = ""
            Select Case HHW.在庫管理区分
                Case "2"
                     gStr = Format(HHW.在庫係数, "###") & "%"
'                Case "3"
'                    .TextMatrix(ID, GridCol(10)) = HHW.在庫発注点
                Case "4"
                     gStr = Format(HHW.在庫日数, "#0.0") & "日"
            End Select
            .TextMatrix(ID, GridCol(19)) = gStr
            ' 発注予定数
            .TextMatrix(ID, GridCol(11)) = HHW.発注予定数
            ' 収容数1
            .TextMatrix(ID, GridCol(12)) = HHW.収容数1
            ' 発注ロット1
            .TextMatrix(ID, GridCol(13)) = HHW.発注ロット1
            ' 切替日
            If Trim(HHW.切替日) <> "" Then
               With CisFun
                    .DateE_BefVal = HHW.切替日
                    .DateEdit
                    gStr = .DateE_AfVal
               End With
               ' 切替日
               .TextMatrix(ID, GridCol(14)) = gStr
               ' 切替便
               .TextMatrix(ID, GridCol(15)) = HHW.切替便
            End If
            ' 収容数2
            .TextMatrix(ID, GridCol(16)) = HHW.収容数2
            ' 発注ロット2
            .TextMatrix(ID, GridCol(17)) = HHW.発注ロット2

            If HHW.Day1 <> 0 Or HHW.Day2 <> 0 Or HHW.Day3 <> 0 Or _
               HHW.Day4 <> 0 Or HHW.Day5 <> 0 Then
               .TextMatrix(ID, GridCol(18)) = "日付指定"
               .TextMatrix(ID, GridCol(20)) = Format(HHW.Day1, "##")
               .TextMatrix(ID, GridCol(21)) = Format(HHW.Day2, "##")
               .TextMatrix(ID, GridCol(22)) = Format(HHW.Day3, "##")
               .TextMatrix(ID, GridCol(23)) = Format(HHW.Day4, "##")
               .TextMatrix(ID, GridCol(24)) = Format(HHW.Day5, "##")
            End If
            If HHW.Week1 <> 0 Or HHW.Week2 <> 0 Or HHW.Week3 <> 0 Or _
               HHW.Week4 <> 0 Or HHW.Week5 <> 0 Then
               .TextMatrix(ID, GridCol(18)) = "曜日指定"
               If HHW.Week1 <> 0 Then .TextMatrix(ID, GridCol(20)) = "月"
               If HHW.Week2 <> 0 Then .TextMatrix(ID, GridCol(21)) = "火"
               If HHW.Week3 <> 0 Then .TextMatrix(ID, GridCol(22)) = "水"
               If HHW.Week4 <> 0 Then .TextMatrix(ID, GridCol(23)) = "木"
               If HHW.Week5 <> 0 Then .TextMatrix(ID, GridCol(24)) = "金"
               
            End If
        End With
        Call HHWReadNext
    Loop
    Call HHWClose
    
    Call CisVsGrid.ReNumber
    
    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'******************************************'
'*****     グリッド内容セット(品番情報)*****
'******************************************'
Private Sub HInfoDisp(Hinbn As String, Torcd As String, Ukeir As String)
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long

' 抽出処理 初期ｾｯﾄ
    Call CisVsGridH.Clear
    
    If Hinbn = "" Or Torcd = "" Or Ukeir = "" Then Exit Sub
    
    gSL_Select = "Select HT.取引先,HT.受入,HT.発注比率,ISNULL(TR.略称,'') 取引先名 "
    gSL_Select = gSL_Select & "   From 品番取引先マスタ HT " & vbCrLf
    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TR " & vbCrLf
    gSL_Select = gSL_Select & "     ON  TR.取引先CD = HT.取引先 " & vbCrLf
    gSL_Select = gSL_Select & "     AND TR.取引先区分 = 1 " & vbCrLf
    gSL_Select = gSL_Select & " Where   HT.品番 = '" & Hinbn & "'" & vbCrLf
    gSL_Select = gSL_Select & "     AND HT.品目 = 1 " & vbCrLf
'    gSL_Select = gSL_Select & "     And HT.取引先 = '" & Torcd & "'" & vbCrLf
'    gSL_Select = gSL_Select & "     AND HT.受入  = '" & Ukeir & "'" & vbCrLf
    If Not HTMRead(gSL_Select) Then
       Call HTMClose
       Exit Sub
    End If

    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGridH
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGridH.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGridH.DispMax Then
            .Rows = CisVsGridH.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HTM_RDSTS
        With VsGridH
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            ' 仕入先
            gStr = HTM.取引先
            If Trim(HTM.取引先名) <> "" Then
                gStr = gStr & ":" & HTM.取引先名
            End If
            .TextMatrix(ID, GridCol(1)) = gStr
            ' 受入
            .TextMatrix(ID, GridCol(2)) = HTM.受入
            ' 比率
            If HTM.発注比率 <> 0 Then
               .TextMatrix(ID, GridCol(3)) = Format(HTM.発注比率, "###") & "% "
            End If
       End With
        Call HTMReadNext
    Loop
    Call HTMClose
    With VsGridH
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'***********************************************'
'*****     グリッド内容セット(入荷予定情報)*****
'***********************************************'
Private Sub NInfoDisp(Hinbn As String, Torcd As String, Ukeir As String)
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
    Dim wChuzan     As Long
    
' 抽出処理 初期ｾｯﾄ
    Call CisVsGridY.Clear
    B1lb_Chuzan = ""
    wChuzan = 0
    
    If Hinbn = "" Or Torcd = "" Or Ukeir = "" Then Exit Sub
    
    gSL_Select = " Select HC.納入日,HC.仕入先,HC.受入," & vbCrLf
    gSL_Select = gSL_Select & " SUM(発注数) 発注数," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL(TR.略称,'') 仕入先名" & vbCrLf
    gSL_Select = gSL_Select & "  From 発注テーブル HC" & vbCrLf
    gSL_Select = gSL_Select & "   Left Outer Join 入荷テーブル NK  ( NoLock)" & vbCrLf
    gSL_Select = gSL_Select & "     ON NK.発注管理NO = HC.発注管理NO" & vbCrLf
    gSL_Select = gSL_Select & "   Left Outer Join 品番取引先マスタ HT ( NoLock)" & vbCrLf
    gSL_Select = gSL_Select & "     ON  HT.品番 = HC.品番" & vbCrLf
    gSL_Select = gSL_Select & "     And HT.取引先 = HC.仕入先" & vbCrLf
    gSL_Select = gSL_Select & "     And HT.受入 = HC.受入" & vbCrLf
    gSL_Select = gSL_Select & "   Left Outer Join 取引先マスタ TR ( Nolock )" & vbCrLf
    gSL_Select = gSL_Select & "     ON  TR.取引先区分 = 1" & vbCrLf
    gSL_Select = gSL_Select & "     And TR.取引先CD = HC.仕入先" & vbCrLf
    gSL_Select = gSL_Select & "  Where HC.品番 = '" & Hinbn & "'" & vbCrLf
'    gSL_Select = gSL_Select & "    And HC.仕入先 = '" & Torcd & "'" & vbCrLf
'    gSL_Select = gSL_Select & "    AND HC.受入  = '" & Ukeir & "'" & vbCrLf
    gSL_Select = gSL_Select & "    And ISNULL(NK.入荷日,'') = ''"
    gSL_Select = gSL_Select & "    Group By HC.納入日,HC.仕入先,HC.受入,ISNULL(TR.略称,'')"
    gSL_Select = gSL_Select & "    Order By HC.納入日,HC.仕入先,HC.受入"
        
    If Not HCTRead(gSL_Select) Then
       Call HCTClose
       Exit Sub
    End If

    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGridY
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGridY.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGridY.DispMax Then
            .Rows = CisVsGridY.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HCT_RDSTS
        With VsGridY
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            ' 入荷予定日
            With CisFun
                 .DateE_BefVal = HCT.納入日
                 .DateEdit
                 gStr = .DateE_AfVal
            End With
            .TextMatrix(ID, GridCol(1)) = gStr
            ' 入荷予定数
            .TextMatrix(ID, GridCol(2)) = Format(HCT.発注数, "#,###")
            
            wChuzan = wChuzan + HCT.発注数
            
            ' 仕入先
            gStr = HCT.仕入先
            If Trim(HCT.仕入先名) <> "" Then
                gStr = gStr & ":" & HCT.仕入先名
            End If
            .TextMatrix(ID, GridCol(3)) = gStr
            ' 受入
            .TextMatrix(ID, GridCol(4)) = HCT.受入
       End With
        Call HCTReadNext
    Loop
    Call HCTClose
    With VsGridY
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With

    B1lb_Chuzan = Format(wChuzan, "#,###")
    
End Sub

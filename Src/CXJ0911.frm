VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form CXJ0911 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "受注データ検索"
   ClientHeight    =   8610
   ClientLeft      =   2265
   ClientTop       =   2865
   ClientWidth     =   14295
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
   ScaleHeight     =   8610
   ScaleWidth      =   14295
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1290
      Left            =   105
      ScaleHeight     =   1230
      ScaleWidth      =   9795
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   435
      Width           =   9855
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1140
         Index           =   0
         Left            =   45
         Top             =   45
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   "納入日-便"
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
         cPositionY      =   50
         Begin CisYMD_With_Btn.CisYMDwB iH1_SDate 
            Height          =   345
            Left            =   570
            TabIndex        =   0
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
         Begin CisText_V60.CisText iH1_SBin 
            Height          =   360
            Left            =   2400
            TabIndex        =   1
            Top             =   360
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            cFormat         =   "00"
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
            Left            =   2400
            TabIndex        =   3
            Top             =   720
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            cFormat         =   "00"
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
            Text            =   "00"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_EDate 
            Height          =   345
            Left            =   570
            TabIndex        =   2
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
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   570
         Left            =   4230
         Top             =   45
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "納入先"
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
         cAlingnment     =   1
         cPositionX      =   50
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   810
            TabIndex        =   6
            Top             =   135
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
         Begin Cis3D_v60.CIS3D H1lb_Torcd 
            Height          =   360
            Left            =   1800
            Top             =   135
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
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   570
         Left            =   8640
         Top             =   45
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "受入"
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
         cAlingnment     =   1
         cPositionX      =   50
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   555
            TabIndex        =   7
            Top             =   105
            Width           =   375
            _ExtentX        =   661
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
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   570
         Left            =   4230
         Top             =   615
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "品番"
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
         cAlingnment     =   1
         cPositionX      =   50
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   585
            TabIndex        =   8
            Top             =   105
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   570
         Left            =   8070
         Top             =   615
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "背番号"
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
         cAlingnment     =   1
         cPositionX      =   50
         Begin CisText_V60.CisText iH1_Seban 
            Height          =   375
            Left            =   840
            TabIndex        =   9
            Top             =   105
            Width           =   750
            _ExtentX        =   1323
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
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D9 
         Height          =   1140
         Index           =   2
         Left            =   2925
         Top             =   45
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   "納　番"
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
         Begin CisText_V60.CisText iH1_SNouNo 
            Height          =   360
            Left            =   510
            TabIndex        =   4
            Top             =   300
            Width           =   735
            _ExtentX        =   1296
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
            Left            =   510
            TabIndex        =   5
            Top             =   660
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
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   315
            Left            =   60
            Top             =   315
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   556
            Caption         =   "開始"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   315
            Left            =   75
            Top             =   675
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   556
            Caption         =   "終了"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
      End
   End
   Begin Cis3D_v60.CIS3D HLB_Title 
      Height          =   390
      Left            =   90
      Top             =   15
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   688
      BackColor       =   16761087
      ForeColor       =   0
      Caption         =   "【 受  注  デ  ー  タ  検  索 】"
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
      Left            =   10620
      ScaleHeight     =   615
      ScaleWidth      =   3480
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1035
      Width           =   3540
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   1755
         Top             =   15
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
         Left            =   2610
         Top             =   15
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
         Left            =   900
         Top             =   15
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
         Top             =   15
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
      Height          =   6735
      Left            =   105
      ScaleHeight     =   6675
      ScaleWidth      =   13995
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1770
      Width           =   14055
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CXJ0911.frx":0000
         Height          =   6525
         Left            =   60
         OleObjectBlob   =   "CXJ0911.frx":0015
         TabIndex        =   10
         Top             =   75
         Width           =   13890
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
      Height          =   585
      Left            =   11910
      ScaleHeight     =   525
      ScaleWidth      =   2190
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   435
      Width           =   2250
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
         Left            =   1290
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   150
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
         Left            =   210
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   75
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
Attribute VB_Name = "CXJ0911"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   受注データ参照
'**       フォームID    :   CXJ0911
'**       処理概要      :
'**
'**       作  成  日    :   2005/05/12  By CIS
'**       変  更  日    :   2014/02/04  By CIS WIN7対応
'**       変  更  日    :
'**
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim mSL_OrderA      As String       ' ソート順(Order BY)昇順
    Dim mSL_OrderD      As String       ' ソート順(Order BY)降順
    Dim mSL_SelectW     As String       ' SQL SELECT(Where)条件文
    Dim mSL_Group       As String       ' (Group BY)
    Dim FormCap         As String
    Dim TorWhere        As String
    
    Dim CisFrm          As CISFormContorl
    Dim CisDBGrid       As New CisDBGrid ' TDBGridPro60ｸﾗｽ定義
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
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
    Me.Caption = FormCap
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
'+---------------------+
'+ TrueGrid60 初期設定  +
'+---------------------+
    With CisDBGrid
        Set .DBGridObj = DBGrid1
        .FrmCaption = Me.Caption
        .MaxSort = 7                        ' ﾚｲｱｳﾄｿｰﾄ最大数
        .LayOut = False                     ' ﾚｲｱｳﾄ変更処理 True:有り,False:無し
        .LayOutMax = 5                      ' ﾚｲｱｳﾄ情報最大数
        .LayOutColMax = 10                  ' ﾚｲｱｳﾄ表示可能項目最大数
        .LayOutPath = gIniExe               ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ場所
        .LayOutName = CisFun.WSID & "J0911.Lay"       ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ名
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        .CsvRegKey = gRegKey
        .CsvValue = "CXJ0911"
        .GridInitGet = False                 ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ
        .GridProcGet = False
        .Init
    End With
    Call GridInitSet
    
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
    '-----2014/02/04 EDIT START-----
    gStr = CisDB.ConnectString
    gConnect = "server=" & CisDB.Server & ";database=" & CisDB.DBName & ";User ID=" & CisDB.User & ";password=" & CisDB.PassWord(GetSetting(gRegKey, "Environment", "DrowsSap", ""))
    
    gStr = Replace(gStr, "Use Procedure", "Extended Properties=" & gConnect & ";Use Procedure")
    
    Adodc1.ConnectionString = gStr
    '-----2014/02/04 EDIT END-------
    Call HeadBodyClear("H")
    RV_Rtn = False
    FormAct = Not FormAct
    
'    Call ReturnPress
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
'   Me.Hide
'   If RV_Call = "Z" Then
'      Unload CKK0028
'   Else
'      Me.Hide
'   End If
    Unload CXJ0911
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
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_SDate" Then iH1_SDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_EDate" Then iH1_EDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_SSTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd(0): Exit Sub
    If PB_Look.Tag = "iH1_Seban" Then Call Look_Hinbn(1): Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn(0): Exit Sub
End Sub
'+-----------------------------+
'+       取引先検索             +
'+-----------------------------+
Private Sub Look_SSTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 3
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Torcd = RV_TorRName
    End If
    
    RV_Rtn = False
End Sub
'+-----------------------------+
'+       取引先受入検索         +
'+-----------------------------+
Private Sub Look_UTorcd(KB As Integer)
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_Torcd = iH1_Torcd
    RV_TorKb = 3
    RV_Kojyo = ""
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Torcd = RV_TorRName
        iH1_Ukeir = RV_Ukeir
    End If

    RV_Rtn = False
End Sub
'+-----------------------------+
'+       品番検索              +
'+-----------------------------+
Private Sub Look_Hinbn(KB As Integer)
    RV_Left = 0
    RV_Top = 0
        
    RVI_Hinmoku = 0 ' 発注
    RVI_SK = 9      '
    RV_Call = ""
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        Select Case KB
            Case 1: iH1_Seban = RV_Seban
            Case Else
                iH1_Hinbn = RV_Hinbn
        End Select
    End If

    RV_Rtn = False
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
    If TorNmGet(iH1_Torcd, 3) Then
        H1lb_Torcd = TRM_RName
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
'+  iH1_Seban
'+-----------------------------+
Private Sub iH1_Seban_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Seban_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
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
       If Not Head1Chk Then GoTo ReturnPress_Ed
       If Not BodySet Then GoTo ReturnPress_Ed
       ProcHB = "B1"
       PB_ENT.BtnCaption = "選択"
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
    
    ' 日付
    If Trim(iH1_SDate) <> "" Then
       If CisFun.Val2(iH1_SDate.cDay) = 0 Then
          iH1_EDate = ""
          iH1_EBin = ""
       End If
    End If
    If Trim(iH1_SDate) <> "" And Trim(iH1_EDate) <> "" Then
       If iH1_SDate & Format(iH1_SBin, "00") > iH1_EDate & Format(iH1_EBin, "00") Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    納入日・便を正しく入力して下さい         "
                .MB_MSG(4) = "        開始　＞   終了             "
                .MB_Button = Error
                .MBOX
            End With
            iH1_SDate.SetFocus
            Exit Function
       End If
    End If
    ' 納入番号
    If InStr(1, iH1_SNouNo, "*") > 0 Then
       iH1_ENouNo = ""
    End If
    
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
    If Adodc1.Recordset.RecordCount <> 0 Then Call BodySet
    SendKeys "{TAB}"
End Sub
'****************************
'*      ボディ内容セット
'****************************
Private Function BodySet() As Boolean
    Dim DateStr As String
    Dim DateSt  As String
    Dim DateEd  As String
    Dim BinSt   As String
    Dim BinEd   As String
    
    
    BodySet = False
    
'       表示条件 セット
    CisDBGrid.Select_Where = ""
    With CisFun
         .Where_CreateKB = NewDocument
         ' 日付
         DateStr = "納入"
         DateSt = "": DateEd = "": BinSt = "": BinEd = ""
         ' 日付 月指定
         If Trim(iH1_SDate) <> "" And CisFun.Val2(iH1_SDate.cDay) = 0 Then
            .Where_Create AlphaMe, "Substring(" & DateStr & "日,1,6)", iH1_SDate
            .Where_Create AlphaMe, DateStr & "便", Format(iH1_SBin, "00")
         End If
         ' 日付開始編集
         If Trim(iH1_SDate) <> "" Then
            DateSt = iH1_SDate
         End If
         If CisFun.Val2(iH1_SBin) <> 0 Then
            BinSt = Format(iH1_SBin, "00")
         End If
         ' 日付終了編集
         If Trim(iH1_EDate) <> "" Then
            DateEd = iH1_EDate
         End If
         If CisFun.Val2(iH1_EBin) <> 0 Then
            BinEd = Format(iH1_EBin, "00")
         End If
         ' 日付開始のみ
         If Trim(DateSt) <> "" And Trim(DateEd) = "" And CisFun.Val2(iH1_SDate.cDay) <> 0 Then
            If Trim(BinSt) = "" Then
               .Where_Create AlphaMe, DateStr & "日", DateSt
            Else
               .Where_Create AlphaMe, DateStr & "日+" & DateStr & "便", DateSt & BinSt
            End If
         End If
         ' 日付終了のみ
         If Trim(DateSt) = "" And Trim(DateEd) <> "" Then
            If Trim(BinEd) = "" Then
               .Where_Create AlphaMe, DateStr & "日", DateEd, 以下
            Else
               .Where_Create AlphaMe, DateStr & "日+" & DateStr & "便", DateEd & BinEd, 以下
            End If
         End If
         ' 日付開始・終了
         If Trim(DateSt) <> "" And Trim(DateEd) <> "" Then
            If Trim(BinSt) = "" And Trim(BinEd) = "" Then
               .Where_Create AlphaMe, DateStr & "日", DateSt, Between, , DateEd
            Else
               If Trim(BinSt) = "" Then BinSt = "00"
               If Trim(BinEd) = "" Then BinEd = "99"
               .Where_Create AlphaMe, DateStr & "日+" & DateStr & "便", DateSt & BinSt, Between, , DateEd & BinEd
            End If
         End If
         ' 納入先
         .Where_Create AlphaMe, "納入先", iH1_Torcd
         ' 受入
         .Where_Create AlphaMe, "受入", iH1_Ukeir
         ' 品番
         .Where_Create AlphaMe, "品番", iH1_Hinbn
         ' 背番号
         .Where_Create AlphaMe, "背番号", iH1_Seban
         ' 納番
         If Trim(iH1_SNouNo) <> "" And Trim(iH1_ENouNo) <> "" Then
            .Where_Create AlphaMe, "納入番号", iH1_SNouNo, Between, , iH1_ENouNo
         Else
            If Trim(iH1_SNouNo) <> "" Then
               .Where_Create AlphaMe, "納入番号", iH1_SNouNo
            End If
            If Trim(iH1_ENouNo) <> "" Then
               .Where_Create AlphaMe, "納入番号", iH1_ENouNo, 以下
            End If
         End If
         
         CisDBGrid.Select_Where = .Where_Phrase
         
    End With
    
' ﾚｲｱｳﾄ表示内容よりﾃﾞｰﾀ連結名を求める
    CisDBGrid.LayOutGet
    
    gSL_Select = "select " & CisDBGrid.LayOutfields
    gSL_Select = gSL_Select & " from 受注出荷状況問合ビュー "

' 分類内容(Order By句) ｾｯﾄ
    With CisDBGrid
        .OrderSet (Op_Sort(0).Value)
        If .Select_OrderBy = "" Then
            BodySet = False
            Exit Function
        End If
    End With
' 抽出文　作成
    gSL_Select = gSL_Select & CisDBGrid.Select_Where & CisDBGrid.Select_OrderBy

    Adodc1.RecordSource = gSL_Select
    DBGrid1.ReBind
    Adodc1.Refresh
    DBGrid1.Refresh
    If Adodc1.Recordset.RecordCount = 0 Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(1) = "    該当するデータはありません  "
            .MB_Button = OK
            .MBOX
        End With
        Exit Function
    End If
    Call TotalDisp
    BodySet = True
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
         Call .GridInitSet(True, , "納入日便", "納 入 日 - 便", 15)
         Call .GridInitSet(True, , "編集納入先", "納入先", 25)
         Call .GridInitSet(True, , "受入", "受", 3)
         Call .GridInitSet(True, , "表示納入番号", "納  番", 9)
         Call .GridInitSet(True, , "表示品番", "品       番", 25)
         Call .GridInitSet(True, , "表示背番号", "背番号", 6)
         Call .GridInitSet(, , "収容数", "収容数", 6)
         Call .GridInitSet(, , "受注枚数", "受注枚数", 8, True)
         Call .GridInitSet(, , "納入数", "受注数", 10, True)
         Call .GridInitSet(True, , "出荷指示書NO", "出荷指示書No", 9)
    End With
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
    If CisDBGrid.ColMoveFlg Then  ' 再表示(分類)
        CisDBGrid.ColMoveFlg = False
        Call BodySet
    End If
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseUp  イベント   》                +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    SendKeys "{RIGHT}"
    If DBGrid1.RowContaining(Y) < 0 Then
       Exit Sub
    End If
    If DBGrid1.ColContaining(x) < 0 Then Exit Sub
    DBGrid1.Row = DBGrid1.RowContaining(Y)
    DoEvents
    Call RtnSet
End Sub
'+-------------------------------------+
'+      選択内容ｾｯﾄ & 復帰ｽﾃｰﾀｽｾｯﾄ      +
'+-------------------------------------+
Private Sub RtnSet()
    With Adodc1
        RV_Long = .Recordset("出荷指示書NO")
    End With
    Adodc1.Recordset.Close
    RV_Rtn = True
    Call PB_END_Click
End Sub

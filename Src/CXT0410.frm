VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXT0410 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "TNS送信データ作成状況問合せ"
   ClientHeight    =   11145
   ClientLeft      =   3990
   ClientTop       =   2895
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11145
   ScaleWidth      =   15105
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 TNS送信データ作成状況問合せ 】"
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   11940
         Top             =   60
         Width           =   2520
         _ExtentX        =   4445
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14520
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
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
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   660
      Top             =   30
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
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
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5610
      ScaleHeight     =   615
      ScaleWidth      =   4155
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   480
      Width           =   4215
      Begin VB.Frame Frm_Sort 
         BackColor       =   &H00808000&
         Caption         =   "《分類》"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   60
         TabIndex        =   14
         Top             =   30
         Width           =   1635
         Begin VB.OptionButton Op_Sort 
            BackColor       =   &H00808000&
            Caption         =   "昇順"
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   180
            Width           =   675
         End
         Begin VB.OptionButton Op_Sort 
            BackColor       =   &H00808000&
            Caption         =   "降順"
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   1
            Left            =   870
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   180
            Value           =   -1  'True
            Width           =   675
         End
      End
      Begin Cis3D_v60.CIS3D PNL_Kensu 
         Height          =   465
         Left            =   1800
         Top             =   90
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   820
         ForeColor       =   192
         Caption         =   "件数"
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
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cBoderWidth     =   3
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D Lb_Kensu 
            Height          =   225
            Left            =   630
            Top             =   120
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   397
            Caption         =   "99,999/99,999"
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
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8730
      Left            =   165
      ScaleHeight     =   8670
      ScaleWidth      =   14685
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2310
      Width           =   14745
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CXT0410.frx":0000
         Height          =   8520
         Left            =   150
         OleObjectBlob   =   "CXT0410.frx":0015
         TabIndex        =   7
         Top             =   90
         Width           =   14460
      End
   End
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'ｵﾌ固定
      Left            =   450
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "D"
      Top             =   60
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   30
      Top             =   30
   End
   Begin VB.PictureBox H1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   180
      ScaleHeight     =   1035
      ScaleWidth      =   11235
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1170
      Width           =   11295
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   915
         Left            =   6375
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "ﾌｧｲﾙｺｰﾄﾞ"
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
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_FileCd 
            Height          =   375
            Left            =   270
            TabIndex        =   3
            Top             =   405
            Width           =   540
            _ExtentX        =   953
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
            Text            =   "XXX"
            MaxLength       =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   915
         Left            =   7500
         Top             =   60
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "表　　示　　内   容"
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
         cPositionY      =   60
         Begin VB.OptionButton H1Op_DspKB 
            BackColor       =   &H00C0C0C0&
            Caption         =   "正常終了"
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   2310
            TabIndex        =   6
            Top             =   450
            Width           =   1275
         End
         Begin VB.OptionButton H1Op_DspKB 
            BackColor       =   &H00C0C0C0&
            Caption         =   "異常終了"
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   945
            TabIndex        =   5
            Top             =   450
            Width           =   1275
         End
         Begin VB.OptionButton H1Op_DspKB 
            BackColor       =   &H00C0FFC0&
            Caption         =   "全て"
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   4
            Top             =   450
            Width           =   795
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   915
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1614
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
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   6
         Begin CisYMD_With_Btn.CisYMDwB iH1_SDate 
            Height          =   330
            Left            =   930
            TabIndex        =   0
            Top             =   60
            Width           =   1695
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
            cChkResult      =   0   'False
            cCalenderPosition=   2
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
            Height          =   315
            Left            =   420
            Top             =   75
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            ForeColor       =   16711680
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D Disp_Date 
            Height          =   765
            Left            =   60
            Tag             =   "0"
            Top             =   30
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   1349
            ForeColor       =   16711680
            Caption         =   "処理日"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   2
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_EDate 
            Height          =   330
            Left            =   930
            TabIndex        =   1
            Top             =   450
            Width           =   1695
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
            cChkResult      =   0   'False
            cCalenderPosition=   2
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
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   315
            Left            =   420
            Top             =   450
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            ForeColor       =   16711680
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
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   915
         Left            =   2775
         Top             =   60
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "契　約　先(送信先)"
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
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_KTorcd 
            Height          =   360
            Left            =   90
            TabIndex        =   2
            Top             =   420
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
         Begin Cis3D_v60.CIS3D H1lb_KTorNm 
            Height          =   360
            Left            =   1080
            Top             =   420
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXX"
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
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   10500
      ScaleHeight     =   645
      ScaleWidth      =   4395
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   450
      Width           =   4455
      Begin CisBtn_60.CisBtn PB_Prt 
         Height          =   585
         Left            =   900
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
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   3510
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
         Left            =   2640
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
         Left            =   1770
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   30
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
   Begin VB.Menu pop_Menu 
      Caption         =   "POP"
      Visible         =   0   'False
      Begin VB.Menu Pop_ReSend 
         Caption         =   "再送信"
      End
      Begin VB.Menu Pop_ReCrt 
         Caption         =   "再処理(作成)"
      End
      Begin VB.Menu Pop_Send 
         Caption         =   "再処理(作成・送信)"
      End
      Begin VB.Menu Pop_Print020 
         Caption         =   "エラー印刷"
      End
   End
End
Attribute VB_Name = "CXT0410"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   ＴＮＳ送信データ作成状況問合せ
'**       フォームID    :   CXT0410
'**       処理概要      :
'**
'**       作  成  日    :   2006/02/13  By CIS
'**       変  更  日    :   2006/03/09  By CIS ファミック仕様にカスタマイズ
'**       変  更  日    :   2007/03/09  By CIS 品番チェック機能追加
'**                                     R07(TNS)対応(トヨタ紡織)
'**                                     ファミック仕様削除
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim MaxCnt          As Long
    Dim NowCnt          As Long
    Dim mSL_SelectW     As String       ' SQL SELECT(Where)条件文
    
    Dim CisFrm          As CISFormContorl
    
    Dim mGridMoveCol    As Boolean
    
    Dim mOPGotColor     As Long
    Dim mOPLostColor    As Long
    
    Dim ErrMsgFlg       As Boolean
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
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
'
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始

    mOPGotColor = H1Op_DspKB(0).BackColor
    mOPLostColor = H1Op_DspKB(1).BackColor

    Pop_ReSend.Visible = False
    Pop_Send.Visible = False
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
    H1Op_DspKB(0).Value = True
    Op_Sort(1).Value = True
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
'
' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    
    gSL_Select = "SELECT Max(開始日時) 開始日時 From TNS送信履歴テーブル "
    If TSRRead(gSL_Select, 1) Then
        If Format(TSR.開始日時, "yyyymmdd") > "20000101" Then
           iH1_SDate = Format(TSR.開始日時, "yyyymmdd")
        End If
    End If
    If iH1_SDate = "" Then
       iH1_SDate = Format(Now(), "yyyymmdd")
    End If
    FormAct = Not FormAct
'
    DoEvents
'   #-----------------------------#
'   #     TNS送信　初期処理       #
'   #-----------------------------#
    Ts_ProcType = 1
    Call TnsSend_Init(False)

'通信ﾊﾟｯｹｰｰｼﾞｴﾗｰﾒｯｾｰｼﾞ更新
'' 通信ﾊﾟｯｹｰｼﾞｴﾗｰﾌｧｲﾙ名
'    With CisFun
'        .INI_Section = "TNS送信"
'        .INI_Default = ""
'        .INI_Key = "EDIErrMsg"
'        .GetIni
'        If Trim(.INI_String) = "" Then
'           .INI_String = "C:\EDI\Sys\Dat\Mst\"
'           .PutIni
'        End If
'        Ts_ErrFName = Trim(.INI_String)
'    End With
'    If Right(Ts_ErrFName, 1) <> "\" Then
'       Ts_ErrFName = Ts_ErrFName & "\"
'    End If
'    Ts_ErrFName = Ts_ErrFName & "Err_info.mst"
'    ErrMsgFlg = False
'    If Dir(Ts_ErrFName) <> "" Then
'        gStr = FileDateTime(Ts_ErrFName)
'
'        Call ItemsClearTET
'        With TET
'            .ERRCODE = "ErrKey"
'            If TETRead("", 1) Then
'                If Trim(.ERRMSG1) <> gStr Then
'                    CisDB.SQL = "DELETE FROM TNSERRMSG   "
'                    Call CisDB.DBExec
'
'                    .ERRMSG1 = gStr
'                    ErrMsgFlg = True
'                End If
'            Else
'                .ERRMSG1 = gStr
'                ErrMsgFlg = True
'            End If
'        End With
'        If ErrMsgFlg Then
'           Call CisFrm.MousePT(11)
'           frmErrMsg.Show vbModal
'           Unload frmErrMsg
'           Set frmErrMsg = Nothing
'           Call CisFrm.MousePT(1)
'        End If
'    End If
    
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
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
           Case vbKeyF8:       If PB_Prt.Visible Then Call PB_Prt_Click    '【印刷】    '-- 2007/03/09 Insert
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
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call DBGridClr
      iH1_SDate.SetFocus
'      GoTo PB_CAN_Ed
   End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'------------------------------------------ 2007/03/09 Insert
'+----------------------------+
'+    印  刷  キ ー(F8)       +
'+----------------------------+
Private Sub PB_Prt_Click()
    If Adodc1.Recordset Is Nothing Then Exit Sub
    If Adodc1.RecordSource = "" Then Exit Sub
    
    TSR.履歴管理NO = Trim(Adodc1.Recordset("送信管理NO"))

    Call C04_ErrList(True)

End Sub
'------------------------------------------ 2007/03/09 Insert End

'+----------------------------+
'+    検　索  キ ー(F7)        +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_SDate" Then Call iH1_SDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_EDate" Then Call iH1_EDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_KTorcd" Then Call Look_KTorcd: Exit Sub
End Sub
'+----------------------------+
'+      送信先(契約先)検索     +
'+----------------------------+
Private Sub Look_KTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 5
    RV_Call = ""
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_KTorcd = RV_TorcdK
        H1lb_KTorNm = RV_TorRName
    End If
End Sub
'+------------------------------------+
'+  iH1_SDate(Got_Focus)
'+------------------------------------+
Private Sub iH1_SDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
'+------------------------------------+
'+  iH1_SDate(Lost_Focus)
'+------------------------------------+
Private Sub iH1_SDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'+------------------------------------+
'+  iH1_EDate(Got_Focus)
'+------------------------------------+
Private Sub iH1_EDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
'+------------------------------------+
'+  iH1_EDate(Lost_Focus)
'+------------------------------------+
Private Sub iH1_EDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'+-----------------------------+
'+  iH1_KTorcd
'+-----------------------------+
Private Sub iH1_KTorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name:       PB_Look.Visible = True
End Sub
Private Sub iH1_KTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    H1lb_KTorNm = ""
    If TorNmGet(iH1_KTorcd, 5) Then
       H1lb_KTorNm = TRM_RName
    End If
End Sub
'****************************
'*                          *
'*      HEAD/BODY 処理      *
'*                          *
'****************************
'****************************
'*      入力切替 処理       *
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
          PB_ENT.BtnCaption = "開始"
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
          PB_ENT.BtnCaption = "表示"
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    PB_CAN.Visible = B1Mode
    PB_Look.Visible = H1Mode
    PB_ENT.Visible = H1Mode
    '-------------------------------------- 2007/03/09 Insert
    PB_Prt.Visible = False
    '-------------------------------------- 2007/03/09 Insert
    If ProcHB = "H1" Then
       Lb_Kensu = ""
       PB_Look.Visible = False
    End If
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
    If Trim(iH1_SDate.cDay) <> "" Then
        If Not iH1_SDate.cChkResult Then
           With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    正しい日付にて入力してください（開始）"
                .MB_Button = Error
                .MBOX
           End With
           iH1_SDate.SetFocus
           Exit Function
        End If
    End If
    
    If Trim(iH1_EDate.cDay) <> "" Then
        If Not iH1_EDate.cChkResult Then
           With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    正しい日付にて入力してください（終了）"
                .MB_Button = Error
                .MBOX
           End With
           iH1_EDate.SetFocus
           Exit Function
        End If
        If iH1_SDate > iH1_EDate Then
           With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付を正しく入力してください（開始＞終了）"
                .MB_Button = Error
                .MBOX
           End With
           iH1_EDate.SetFocus
           Exit Function
        End If
    End If
    
    If Not DBInput Then
       With CisFun
             .MB_Lines = 4
             .MB_MSG(1) = "    該当データが存在しません｡            "
             .MB_MSG(3) = "  再度抽出条件を入力してください            "
             .MB_Button = Error
             .MBOX
       End With
       Call DispChange(ProcHB)
       Exit Function
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
''******************************************'
''*****         表  示  内  容        *****
''******************************************'
Private Sub H1Op_DspKB_Click(Index As Integer)
    H1Op_DspKB(0).Tag = Index
    For Each gObj In H1Op_DspKB
        If gObj.Index = Index Then
           gObj.BackColor = mOPGotColor
        Else
           gObj.BackColor = mOPLostColor
        End If
    Next gObj
End Sub
''******************************************'
''*****          分  類  表  示         *****
''******************************************'
Private Sub Op_Sort_Click(Index As Integer)
    For Each gObj In Op_Sort
        If gObj.Index = Index Then
           gObj.BackColor = mOPGotColor
           gObj.ForeColor = &H80000012
        Else
           gObj.BackColor = Frm_Sort.BackColor
           gObj.ForeColor = &HFFFFFF
        End If
    Next gObj
    If ProcHB = "H1" Then Exit Sub
    Call DBInput
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    DoEvents
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
       Call ErrDisp
    End If
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub

'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(Optional InType As String = "") As Boolean
    Dim wDate   As String

    DBInput = True
    On Error GoTo DBInput_Err:
    DoEvents
    If ProcHB = "H1" And InType = "" Then
'       表示条件 セット
'       入力 Select 文作成
        mSL_SelectW = ""
'       日付
        With CisFun
            .Where_CreateKB = NewDocument
            If Trim(iH1_SDate) = "" And Trim(iH1_EDate) = "" Then
                ' 全件対象
            Else
               If Trim(iH1_SDate) <> "" And Trim(iH1_SDate.cDay) = "" Then
                  iH1_EDate = ""
               End If
               If Trim(iH1_SDate.cDay) = "" Then
                   If Trim(iH1_EDate) <> "" Then
                      ' 終了 年月日
                      .Where_Create AlphaMe, "CONVERT(CHAR(8),開始日時,112)", iH1_EDate, 以下
                   Else
                       ' 開始　年月
                      .Where_Create AlphaMe, "CONVERT(CHAR(6),開始日時,112)", Mid(iH1_SDate, 1, 6)
                   End If
               Else
                   If Trim(iH1_EDate) = "" Then
                      ' 開始 年月日
                      .Where_Create AlphaMe, "CONVERT(CHAR(8),開始日時,112)", iH1_SDate
                   Else
                       ' 開始・終了年月日
                      .Where_Create AlphaMe, "CONVERT(CHAR(8),開始日時,112)", iH1_SDate, 以上
                      .Where_Create AlphaMe, "CONVERT(CHAR(8),開始日時,112)", iH1_EDate, 以下
                   End If
               End If
            End If
            .Where_Create AlphaMe, "TR.送信先", iH1_KTorcd
            .Where_Create AlphaMe, "TR.FCODE", iH1_FileCd
            mSL_SelectW = mSL_SelectW & .Where_Phrase
        End With
        If Not H1Op_DspKB(0) Then
            If H1Op_DspKB(1) Then
                If mSL_SelectW = "" Then
                    mSL_SelectW = " WHERE TR.処理コード <> 0 "
                Else
                    mSL_SelectW = mSL_SelectW & " AND TR.処理コード <> 0 "
                End If
            Else
                If mSL_SelectW = "" Then
                    mSL_SelectW = " WHERE TR.処理コード = 0 "
                Else
                    mSL_SelectW = mSL_SelectW & " AND TR.処理コード = 0 "
                End If
            End If
        End If
'
        
    End If
    
    gSL_Select = "Select TR.履歴管理NO,TR.開始日時,TR.FCODE,ISNULL(TS.ファイル名称,'') F名称,"
    gSL_Select = gSL_Select & "    CASE WHEN TR.実行区分 = 0 THEN '作成'"
    gSL_Select = gSL_Select & "             WHEN TR.実行区分 = 1 THEN '送信'"
    gSL_Select = gSL_Select & "             WHEN TR.実行区分 = 2 THEN '再送信'"
    gSL_Select = gSL_Select & "             WHEN TR.実行区分 = 3 THEN '再処理(送信)'"
    gSL_Select = gSL_Select & "             WHEN TR.実行区分 = 4 THEN '再処理(作成)'"
    gSL_Select = gSL_Select & "            Else '' END 処理,"
    gSL_Select = gSL_Select & "    TR.データ件数,TR.送信件数,TR.送信連番,TR.実行区分,ISNULL(TS.送信,0) 送信,"
    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(SY.値名称,'') = '' THEN 'その他エラー(' + TR.処理結果 + ')'"
    gSL_Select = gSL_Select & "                   ELSE ISNULL(SY.値名称,'') END 処理結果,"
    gSL_Select = gSL_Select & "    TR.ERR,TR.終了日時,TR.対象ファイル,"
    gSL_Select = gSL_Select & "    ISNULL(TS.ゼロ件データ,0) ゼロ件データ,TR.データ件数,"
    gSL_Select = gSL_Select & "    ISNULL(TM.略称,'') 送信先名,TR.送信先,ISNULL(TM2.略称,'') 契約先名,TR.契約先,TR.送信先区分,"
    gSL_Select = gSL_Select & "    TS.識別子,TS.TNS宛先,TR.送信管理NO,TR.自動区分,"
    '----------------------------------------------------------------------- 2007/03/09 Delete
'    gSL_Select = gSL_Select & "    NT.開始連番 , NT.終了連番 "              '2006/03/09
    '----------------------------------------------------------------------- 2007/03/09 Delete End
    '----------------------------------------------------------------------- 2007/03/09 Insert
    gSL_Select = gSL_Select & "    TR.ERR件数 "
    '----------------------------------------------------------------------- 2007/03/09 Insert End
    gSL_Select = gSL_Select & "    From TNS送信履歴テーブル TR"
    gSL_Select = gSL_Select & "    Left Outer Join TNS送信設定テーブル TS"
    gSL_Select = gSL_Select & "      ON  TS.送信先 = TR.契約先"
    gSL_Select = gSL_Select & "      And TS.FCODE = TR.FCODE"
    gSL_Select = gSL_Select & "      AND TS.代行送信先 = CASE WHEN TR.送信先 = TR.契約先 THEN '' ELSE TR.送信先 END"
    gSL_Select = gSL_Select & "    Left Outer Join 名称マスタ SY"
    gSL_Select = gSL_Select & "      ON  SY.区分名称 = '送受信処理結果'"
    gSL_Select = gSL_Select & "      And SY.値 = TR.処理結果"
    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TM"
    gSL_Select = gSL_Select & "      ON  TM.取引先CD = TR.送信先"
    gSL_Select = gSL_Select & "      AND TM.取引先区分 = 1"
    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TM2"
    gSL_Select = gSL_Select & "      ON  TM2.取引先CD = TR.契約先"
    gSL_Select = gSL_Select & "      AND TM2.取引先区分 = 1"
''    '[][][][][][][][][][][][][ 2006/03/09 ][][][][][][][][][][][][][][][][][][][][][][][][][][][][][]
''    gSL_Select = gSL_Select & "    Left Outer Join ( SELECT TNS作成管理NO,MIN(TNS連番) 開始連番,"
''    gSL_Select = gSL_Select & "                                           MAX(TNS連番) 終了連番 "
''    gSL_Select = gSL_Select & "                      FROM 入荷テーブル GROUP BY TNS作成管理NO ) NT"
''    gSL_Select = gSL_Select & "      ON  NT.TNS作成管理NO = TR.送信管理NO "
''    '[][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][]
    
    gSL_Select = gSL_Select & mSL_SelectW
    gSL_Select = gSL_Select & " Order By 開始日時 "

' 分類内容(Order By句) ｾｯﾄ
    If Op_Sort(1).Value = True Then
        gSL_Select = gSL_Select & " Desc "
    End If
    If Not DBGridDisp Then
       DBInput = False
       GoTo DBInput_Ed
    End If
    
    On Error Resume Next
    MaxCnt = 0
    If Not Adodc1.Recordset Is Nothing Then
       MaxCnt = Adodc1.Recordset.RecordCount
    End If
    If MaxCnt = 0 Then
       DBInput = False
       GoTo DBInput_Ed
    End If
    Lb_Kensu.Visible = True
    Call TotalDisp
    On Error GoTo 0
DBInput_Ed:
    Call CisFrm.MousePT(0)
    SendKeys "{END}"
    SendKeys "{RIGHT}"
    SendKeys "{LEFT}"
    SendKeys "{HOME}"
    Exit Function
DBInput_Err:
    CisFun.ErrorBox
    Call CisFrm.MousePT(0)
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      メニュー処理(PopUp)                                  +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+      再送信
'+----------------------------------------------------------+
Private Sub Pop_ReSend_Click()
    If Not ReSendCheck Then Exit Sub
    
    Ts_TNSFName = Trim(Adodc1.Recordset("対象ファイル"))
    If Dir(Ts_TNSFName) = "" Then
       With CisFun
            .MB_Lines = 6
            .MB_MSG(1) = "   　　　対象ファイル未登録"
            .MB_MSG(3) = "     (" & Trim(Adodc1.Recordset("F名称")) & ")     "
            .MB_MSG(5) = "  ﾌｧｲﾙ名 = " & Trim(Adodc1.Recordset("対象ファイル"))
            .MB_Title = "TNS再送信"
            .MB_Button = OK
            .MBOX
       End With
       Exit Sub
    End If
    
    Call Get_TSRNumber(Trim(Adodc1.Recordset("送信先")), Trim(Adodc1.Recordset("FCODE")))
    
    gLong = Adodc1.Recordset("履歴管理NO")
    gSL_Select = "Select * From TNS送信履歴テーブル "
    gSL_Select = gSL_Select & " Where 履歴管理NO = " & gLong
    
    If Not TSRRead(gSL_Select, 1) Then Exit Sub
    With CisFun
         .MB_Lines = 6
         .MB_MSG(1) = "   　　　再送信処理を行いますか？"
         .MB_MSG(3) = "         【 " & Trim(Adodc1.Recordset("送信先名")) & "  】         (" & Trim(Adodc1.Recordset("F名称")) & ")     "
         .MB_MSG(5) = "  ﾌｧｲﾙ名 = " & Trim(Adodc1.Recordset("対象ファイル"))
         .MB_Title = "TNS再送信"
         .MB_Button = OK_CAN
         If Not .MBOX Then Exit Sub
    End With
    
    Call TNSSend_LogPut("S")
    
    With TSR
         .履歴管理NO = TSR_RNo + 1
         .実行区分 = 2
         .開始日時 = Now()
         .ERR = ""
    End With
    Call TSRInsert
    
    TSR.実行区分 = 2
    TSR.送信開始日時 = Now()
    Call Middle_TSRUpDate("14", "TNS再送信エラー")
    
    gStr = "【" & Trim(Adodc1.Recordset("F名称")) & "】 処理開始 "
    Call TNSSend_MsgDisp(gStr, gStr)
    gStr = "送信先 = " & TSR.送信先 & Trim(Adodc1.Recordset("送信先名")) & "  ﾌｧｲﾙｺｰﾄﾞ = " & TSR.FCODE
    Call TNSSend_MsgDisp("", gStr, 20)
    
    gStr = "【　　TNS送信処理　開始(再送信)　　】"
    Call TNSSend_MsgDisp("", gStr, 10)
               
    Ts_Shikibetsu = RTrim(Adodc1.Recordset("識別子"))
    Ts_Atesaki = RTrim(Adodc1.Recordset("TNS宛先"))
               
    gStr = "送信元ﾌｧｲﾙ名=" & Ts_TNSFName
    Call TNSSend_MsgDisp("", gStr, 15)
    gStr = "識別子=" & Ts_Shikibetsu & "/TNS宛先=" & Ts_Atesaki
    Call TNSSend_MsgDisp("", gStr, 15)
               
    Ts_SendProc = True
    If Not TNS_File_Send_Api Then
       Ts_SendProc = False
       gStr = "【　　TNS送信処理　異常終了　　】ｴﾗｰｺｰﾄﾞ = " & Ts_TNSReturn
       Call TNSSend_MsgDisp("", gStr, 10)
       TSR.ERR = Ts_TNSReturn
    Else
       gStr = "【　　TNS送信処理　終了　　】"
       Call TNSSend_MsgDisp("", gStr, 10)
    End If
    
'
    TSR.送信終了日時 = Now()
    Call Middle_TSRUpDate("UPD")  ' 履歴更新
    
    '+-----------> 終了処理
    If Ts_SendProc Then
       Call Middle_TSRUpDate("00")   ' 正常終了
    End If
    
    Call TNSSend_LogPut("E")
    
    Call DBInput
    
End Sub
'+--------------------------------+
'+      再処理 再作成 & 再送信
'+--------------------------------+
Private Sub Pop_Send_Click()
    Ts_ProcType = 3
    Call CallProc
End Sub
'+--------------------------------+
'+      再処理 再作成
'+--------------------------------+
Private Sub Pop_ReCrt_Click()
    Ts_ProcType = 4
    Call CallProc
End Sub
'-------------------------------------------------------- 2007/03/09 Insert
'+--------------------------------+
'+      エラーリスト印刷
'+--------------------------------+
Private Sub Pop_Print020_Click()
    
    Call PB_Prt_Click
End Sub
'-------------------------------------------------------- 2007/03/09 Insert End
Private Sub CallProc()
    
    With CisFun
         .MB_Lines = 6
        If Ts_ProcType = 4 Then
             .MB_MSG(1) = "   　　　再処理(作成)を行いますか？                      "
        Else
             .MB_MSG(1) = "   　　　再処理(作成・送信)を行いますか？                "
        End If
         .MB_MSG(4) = "         【 " & Trim(Adodc1.Recordset("送信先名")) & "  】         "
         .MB_MSG(5) = "             (" & Trim(Adodc1.Recordset("F名称")) & ")             "
         .MB_Title = "TNS再処理"
         .MB_Button = OK_CAN
         If Not .MBOX Then Exit Sub
    
         .MB_Lines = 6
         .MB_MSG(1) = "   　　　送信連番(SEQNO)の加算を行いますか？                      "
         .MB_MSG(3) = "         加算する場合　： 今日の最終送信連番＋１"
         .MB_MSG(4) = "         加算しない場合： 今日の最終送信連番"
         .MB_Title = "送信連番確認"
         .MB_Button = OK_CAN
         If .MBOX Then
            Ts_SeqNoAdd = 0
         Else
            Ts_SeqNoAdd = 1
         End If
    End With
    
    Call TNSSend_LogPut("S")
    
    ReDim Ts_SndTorcd(1)
    ReDim Ts_SndFCode(1)
    ReDim Ts_SndKTorcd(1)
    ReDim Ts_SndSendKb(1)
    Ts_SndTorcd(0) = Adodc1.Recordset("送信先")
    Ts_SndFCode(0) = Adodc1.Recordset("FCODE")
    Ts_SndKTorcd(0) = Adodc1.Recordset("契約先")
    Ts_SndSendKb(0) = Adodc1.Recordset("送信先区分")
    
    Ts_ReSendNo = Adodc1.Recordset("送信管理NO")
'
    Ts_ProcErr = False
    
    CXT0020.Show vbModal
    Unload CXT0020
    Set CXT0020 = Nothing
    
    Call DBInput
End Sub

'+--------------------------------+
'+  再送信　対象確認
'+--------------------------------+
Private Function ReSendCheck() As Boolean
    ReSendCheck = False
    
    If Adodc1.Recordset Is Nothing Then Exit Function
    If Adodc1.RecordSource = "" Then Exit Function
    If Adodc1.Recordset.EOF Then Exit Function
    
    If Not Ts_TNSPakegeFlg Then Exit Function
    
''''    gByte = Adodc1.Recordset("実行区分")
''''    If gByte = 0 Then Exit Function
'''
''''    If Adodc1.Recordset("送信") = 0 Then Exit Function
    
    If Adodc1.Recordset("ゼロ件データ") = 1 And Adodc1.Recordset("データ件数") = 0 Then Exit Function
    
    gStr = Trim(Adodc1.Recordset("対象ファイル"))
    
    If gStr = "" Then Exit Function

    ReSendCheck = True
    
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(TrueDBGrid)                             +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+      ADODC
'+----------------------------------------------------------+
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Adodc1.RecordSource <> "" Then
       Call TotalDisp
       PB_ENT.Visible = False
       On Error Resume Next
       If Trim(Adodc1.Recordset("ERR")) = "" Then
       Else
          PB_ENT.Visible = True
       End If
       On Error GoTo 0
       If ReSendCheck Then
          Pop_ReSend.Enabled = True
          Pop_Send.Enabled = True
       Else
          Pop_ReSend.Enabled = False
          Pop_Send.Enabled = False
       End If
    '--------------------------------------------------------------- 2007/03/09 Insert
       Pop_Print020.Enabled = False
       PB_Prt.Visible = False
       If CisFun.Val2(Adodc1.Recordset("ERR件数")) <> 0 Then
          Pop_Print020.Enabled = True
          PB_Prt.Visible = True
       End If
    '--------------------------------------------------------------- 2007/03/09 Insert End

    End If
End Sub
'+----------------------------------------------------------+
'+      ADODC1(件数表示)                                     +
'+----------------------------------------------------------+
Private Sub TotalDisp()
    If Adodc1.Recordset Is Nothing Then Exit Sub
    If Adodc1.RecordSource = "" Then Exit Sub
    NowCnt = Adodc1.Recordset.AbsolutePosition
    MaxCnt = Adodc1.Recordset.RecordCount
    If NowCnt < 0 Then Exit Sub
    Lb_Kensu.Caption = Format(NowCnt, "###,##0") & " / " & _
                       Format(MaxCnt, "###,##0")
                       
End Sub
''+----------------------------------------------------------+
''+      グリッド    《 MouseDown  イベント   》               +
''+----------------------------------------------------------+
'Private Sub DBGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If DBGrid1.ColContaining(X) < 0 Or DBGrid1.RowContaining(Y) < 0 Then Exit Sub
'    DBGrid1.Row = DBGrid1.RowContaining(Y)
'End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseUp  イベント   》                +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If DBGrid1.ColContaining(X) < 0 Then Exit Sub
    DBGrid1.Row = DBGrid1.RowContaining(Y)
' サブメニュー表示
    If DBGrid1.RowContaining(Y) >= 0 Then
        If Button = 2 Then
           PopupMenu pop_Menu
           Exit Sub
        End If
    End If
End Sub
Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace And PB_ENT.Visible Then
    ' ｴﾗｰ内容表示
        Call ErrDisp
    End If
End Sub
'+-----------------------+
'+  グリッド内容初期化    +
'+-----------------------+
Private Sub DBGridClr()
    On Error Resume Next
    Adodc1.Recordset.Close
    Adodc1.RecordSource = ""
    DBGrid1.Close True
    Lb_Kensu.Caption = ""
    On Error GoTo 0
End Sub
'+---------------------+
'+  グリッド内容表示    +
'+---------------------+
Private Function DBGridDisp()
    DBGridDisp = False
    On Error GoTo DBGridDisp_ER:
    If Me.MousePointer <> 11 Then
       Me.MousePointer = 11
    End If
    On Error Resume Next
    Adodc1.RecordSource = gSL_Select
    DBGrid1.ReBind
    Adodc1.Refresh
    DBGrid1.Refresh
    With DBGrid1
         .SetFocus
    End With
    SendKeys "{HOME}"
    If Me.MousePointer <> 1 Then
       Me.MousePointer = 1
    End If
    DBGridDisp = True
    On Error GoTo 0
    Exit Function
DBGridDisp_ER:
    If ERR <> 40002 Then        ' RDO タイムアウト
       CisFun.ErrorBox
    End If
End Function
'+------------------------+
'+      ｴﾗｰ内容表示       +
'+------------------------+
Private Sub ErrDisp()
    Static wErrCode As String
    
    gStr = Trim(Adodc1.Recordset("ERR"))
    If gStr = "" Then Exit Sub
'
    wErrCode = gStr
    gSL_Select = "Select * From TNSERRMSG"
    gSL_Select = gSL_Select & " where ERRCODE = '" & wErrCode & "'"
    If Not TETRead(gSL_Select, 1) Then
       With CisFun
            .MB_Lines = 3
            .MB_MSG(1) = "エラー内容未登録 ErrCode=" & wErrCode
            .MB_Button = OK
            .MBOX
       End With
       Exit Sub
    End If
'
    If Trim(TET.ERRLEBEL) <> "" Then
       With CisFun
            .MB_Lines = 7
            .MB_MSG(1) = "   テキスト変換処理エラー(" & Trim(Adodc1.Recordset("F名称")) & ")     "
            .MB_MSG(3) = "  理由ｺｰﾄﾞ = " & Mid(TET.ERRLEBEL, 1, 3) & _
                         "  内部ｺｰﾄﾞ = " & Mid(TET.ERRLEBEL, 4, 3) & _
                         "  ｴﾗｰﾚﾍﾞﾙ = " & TET.ERRLEBEL & "     "
            .MB_MSG(5) = "  ｴﾗｰ内容  = " & Trim(TET.ERRMSG1) & "      "
            .MB_MSG(7) = "  ｴﾗｰ原因  = " & Trim(TET.ERRMSG2) & "      "
            .MB_Title = "TNS送信処理ｴﾗｰ表示"
            .MB_Button = Error
            .MBOX
       End With
    End If
'
End Sub


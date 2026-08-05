VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Begin VB.Form COY0110 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "売上予想処理"
   ClientHeight    =   11235
   ClientLeft      =   1230
   ClientTop       =   705
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
   ForeColor       =   &H00004000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.Timer Timer3 
      Left            =   900
      Top             =   1035
   End
   Begin VB.PictureBox C1_Area1 
      BackColor       =   &H80000007&
      Height          =   3810
      Left            =   10575
      ScaleHeight     =   3750
      ScaleWidth      =   4035
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4500
      Visible         =   0   'False
      Width           =   4095
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   0
         Left            =   45
         Top             =   1095
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12640511
         ForeColor       =   0
         Caption         =   "IN:型費件数"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Input 
            Height          =   360
            Index           =   2
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   33023
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
      End
      Begin Cis3D_v60.CIS3D PFXLabel1 
         Height          =   405
         Index           =   0
         Left            =   45
         Top             =   30
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   714
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "<<<　処 理 対 象 件 数　>>>"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cFont3D         =   2
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   1
         Left            =   45
         Top             =   570
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12640511
         ForeColor       =   0
         Caption         =   "IN:内示件数"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Input 
            Height          =   360
            Index           =   1
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   33023
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   3
         Left            =   45
         Top             =   1620
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648447
         ForeColor       =   0
         Caption         =   "OUT:単価未決件数"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   1
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   2
         Left            =   45
         Top             =   2670
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "OUT:売上予想件数"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   3
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   4
         Left            =   45
         Top             =   2145
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648447
         ForeColor       =   0
         Caption         =   "OUT:品番未登録数"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   2
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   5
         Left            =   45
         Top             =   3195
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "OUT:売上型費件数"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   4
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00808080&
      Height          =   5235
      Left            =   4110
      ScaleHeight     =   5175
      ScaleWidth      =   7080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3075
      Width           =   7140
      Begin VB.PictureBox Picture2 
         Height          =   5025
         Left            =   60
         ScaleHeight     =   4965
         ScaleWidth      =   6900
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   60
         Width           =   6960
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   4830
            Left            =   90
            Top             =   90
            Width           =   6765
            _ExtentX        =   11933
            _ExtentY        =   8520
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
            cBoderStyle     =   1
            Begin Cis3D_v60.CIS3D UC_3D17 
               Height          =   2415
               Index           =   1
               Left            =   60
               Top             =   45
               Width           =   6645
               _ExtentX        =   11721
               _ExtentY        =   4260
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cPositionX      =   150
               Begin Cis3D_v60.CIS3D CIS3D1 
                  Height          =   330
                  Index           =   1
                  Left            =   495
                  Top             =   510
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   582
                  ForeColor       =   16711680
                  Caption         =   "今回対象年月"
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
                  cAlingnment     =   1
                  cBoderStyle     =   2
               End
               Begin CisYM_V60.CisYM iH1_YM 
                  Height          =   435
                  Left            =   2025
                  TabIndex        =   0
                  Top             =   450
                  Width           =   2025
                  _ExtentX        =   2646
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   14.25
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
               Begin Cis3D_v60.CIS3D CIS3D1 
                  Height          =   330
                  Index           =   2
                  Left            =   585
                  Top             =   1185
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   582
                  ForeColor       =   16711680
                  Caption         =   "契  約  先"
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
                  cAlingnment     =   1
                  cBoderStyle     =   2
               End
               Begin CisText_V60.CisText iH1_Torcd 
                  Height          =   375
                  Left            =   2010
                  TabIndex        =   1
                  Top             =   1185
                  Width           =   1020
                  _ExtentX        =   1799
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
                  Text            =   "XXXXXXX"
                  MaxLength       =   7
               End
               Begin Cis3D_v60.CIS3D H1lb_Torcd 
                  Height          =   360
                  Left            =   3375
                  Top             =   1185
                  Width           =   2535
                  _ExtentX        =   4471
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
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   30
               End
               Begin Cis3D_v60.CIS3D PB_Torcd 
                  Height          =   360
                  Left            =   3030
                  Top             =   1185
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   635
                  ForeColor       =   65280
                  Caption         =   "▼"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  cButton         =   -1  'True
                  cPositionX      =   20
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D9 
               Height          =   2265
               Left            =   60
               Top             =   2505
               Width           =   6630
               _ExtentX        =   11695
               _ExtentY        =   3995
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
               cBoderStyle     =   1
               Begin VB.PictureBox B1_Area 
                  BackColor       =   &H00FFC0FF&
                  Height          =   555
                  Left            =   60
                  ScaleHeight     =   495
                  ScaleWidth      =   6465
                  TabIndex        =   8
                  TabStop         =   0   'False
                  Top             =   1650
                  Width           =   6525
                  Begin Cis3D_v60.CIS3D lb_msg 
                     Height          =   435
                     Left            =   30
                     Top             =   30
                     Width           =   5595
                     _ExtentX        =   9869
                     _ExtentY        =   767
                     BackColor       =   16761087
                     Caption         =   "【 処理中 】しばらくお待ち下さい"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   14.25
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cBoderColor1    =   16761087
                     cBoderColor2    =   16761087
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   7
                     cBoderStyle     =   2
                  End
               End
               Begin VB.PictureBox Picture1 
                  BackColor       =   &H00FFFFC0&
                  Height          =   1545
                  Index           =   1
                  Left            =   480
                  ScaleHeight     =   1485
                  ScaleWidth      =   6045
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   6105
                  Begin VB.Label Label2 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "( 契約先  )"
                     BeginProperty Font 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   11.25
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00004080&
                     Height          =   225
                     Index           =   3
                     Left            =   495
                     TabIndex        =   18
                     Top             =   630
                     Width           =   1170
                  End
                  Begin VB.Label H1lb_KTorcd 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "XXXXXXX：XXXXXXXXX1XXXXXXXXX2"
                     BeginProperty Font 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   11.25
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Left            =   1830
                     TabIndex        =   17
                     Top             =   630
                     Width           =   4155
                  End
                  Begin VB.Label H1lb_ZYM 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "YYYY 年 MM 月度"
                     BeginProperty Font 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   1830
                     TabIndex        =   16
                     Top             =   210
                     Width           =   2025
                  End
                  Begin VB.Label H1lb_ZDate 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "YYYY年 MM月 DD日 HH:NN"
                     BeginProperty Font 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   1830
                     TabIndex        =   15
                     Top             =   1050
                     Width           =   3075
                  End
                  Begin VB.Label Label2 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "(対象年月)"
                     BeginProperty Font 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   11.25
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00004080&
                     Height          =   225
                     Index           =   1
                     Left            =   495
                     TabIndex        =   14
                     Top             =   225
                     Width           =   1140
                  End
                  Begin VB.Label Label2 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "( 処理日 )"
                     BeginProperty Font 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   11.25
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00004080&
                     Height          =   225
                     Index           =   0
                     Left            =   525
                     TabIndex        =   13
                     Top             =   1065
                     Width           =   1080
                  End
               End
               Begin Cis3D_v60.CIS3D UC_3D10 
                  Height          =   1545
                  Left            =   60
                  Top             =   60
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   2725
                  BackColor       =   0
                  ForeColor       =   65535
                  Caption         =   "前回内容"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
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
                  cAlingnment     =   7
                  cPositionX      =   10
                  cPositionY      =   -290
               End
            End
         End
      End
   End
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   1035
   End
   Begin VB.Timer Timer1 
      Left            =   60
      Top             =   1035
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   75
      TabIndex        =   2
      Text            =   "Dummy"
      Top             =   690
      Width           =   180
   End
   Begin VB.PictureBox Picture7 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   0
   End
   Begin Cis3D_v60.CIS3D PB_ENT 
      Height          =   645
      Left            =   8355
      Top             =   8820
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1138
      ForeColor       =   16576
      Caption         =   "処理開始"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
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
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   645
      Left            =   5535
      Top             =   8820
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1138
      ForeColor       =   16711680
      Caption         =   "終  了"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
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
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D PB_OK 
      Height          =   645
      Left            =   6945
      Top             =   8820
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1138
      ForeColor       =   65535
      Caption         =   "確  認"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
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
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   390
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   688
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 売上予想 作成処理 】"
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
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
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
         Height          =   285
         Left            =   11880
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
      Left            =   90
      TabIndex        =   12
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5535
      TabIndex        =   11
      Top             =   9525
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8340
      TabIndex        =   10
      Top             =   9525
      Width           =   1485
   End
   Begin VB.Label PNL_OK 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6945
      TabIndex        =   9
      Top             =   9525
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "COY0110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   売上予想作成処理
'**       フォームID    :   COY0110
'**       処理概要      :
'**
'**       作  成  日    :   2004/10/28  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim op_Zairyo               As Byte
    Dim wCnt1                   As Integer
    Dim wCnt2                   As Integer
    
    Dim wRClr                   As Integer
    Dim wGClr                   As Integer
    Dim wBClr                   As Integer
    
    Dim STime                   As Date
    Dim ETime                   As Date
    Dim ProcTime                As Long
    
    Dim mGotColor               As Long
    Dim mLostColor              As Long
    
    Dim mShime                  As Byte
    Dim mTorcd                  As String
    Dim mZenYm                  As String
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
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer3.Interval = 300
    Timer3.Enabled = False
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
    
''    mGotColor = H1Op_Shime(0).BackColor
''    mLostColor = H1Op_Shime(1).BackColor
'   < シングルプロセスチェック >
    If Not SingleProcess_Check Then PB_END_Click
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   #     ロゴ表示       #
'   #-------------------#
    CisFrm.BGLDisp BG_Logo, 10
    
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
'   < 前回内容表示 >
    Call ProcRegistry_Check
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                   +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'   ( グラデーション )
    With CisFrm
        .Gradation_Direction = Portrait
        .Gradation_Shading = Deepen
        .Gradation 120, 200, 160
    End With

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
        Case vbKeyF7:       If Dummy.Tag <> "" Then Call PB_Look_Click      '【検索】
        Case vbKeyF12:
                            If PB_END.Visible Then
                                Call PB_END_Click                           '【終了】
                            Else
                                If PB_OK.Visible Then Call PB_OK_Click      '【確認】
                            End If
        Case vbKeyReturn
                            If PB_ENT.Visible Then                          '【入力】
                                Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                Call PB_ENT_Click
                                Dummy.Enabled = False
                            End If
           Case Else:       Exit Function
    End Select
    
    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    
    Call TanErr_List                '単価未決リスト
    Call HinErr_List                '品番未登録リスト

    Call SingleProcess_Release
    Call CisDB.DBDISConnect                         ' ﾃﾞｰﾀﾍﾞｰｽ切断
    Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    End
End Sub

'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Call ReturnPress
End Sub
'*----------------------------------*
'*      確　認　終　了　キ　ー      *
'*----------------------------------*
Private Sub PB_OK_Click()
    Call PB_END_Click
End Sub
'+----------------------------+
'+      検         索         +
'+----------------------------+
Private Sub PB_Look_Click()
    If Dummy.Tag = "iH1_Torcd" Then Call PB_Torcd_Click: Exit Sub
End Sub
'+---------------------+
'+    契約先検索        +
'+---------------------+
Private Sub PB_Torcd_Click()

    RV_TorKb = 2    ' 得意先
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
       iH1_Torcd = RV_TorcdK
       Call iH1_Torcd_LostFocus
    End If

End Sub
'+----------------------------+
'+    iH1_Torcd(GotFocus)    +
'+----------------------------+
Private Sub iH1_Torcd_GotFocus()
    Dummy.Tag = ActiveControl.Name
End Sub
'+----------------------------+
'+    iH1_Torcd(LostFocus)    +
'+----------------------------+
Private Sub iH1_Torcd_LostFocus()
    Dummy.Tag = ""
    If TorNmGet(iH1_Torcd, 2) Then
        H1lb_Torcd = TRM_RName
        If Trim(H1lb_Torcd) = "" Then
           H1lb_Torcd = TRM_Name
        End If
    Else
        H1lb_Torcd = ""
    End If
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*       メッセージ表示         *
'*******************************
Private Sub Timer2_Timer()
    CisFrm.MSGRelate
End Sub
'*******************************
'*       メッセージ表示         *
'*******************************
Private Sub Timer3_Timer()
'    wCnt1 = wCnt1 + wCnt2
'    If wCnt1 > 7 Then
'        wCnt1 = 6
'        wCnt2 = -1
'    Else
'        If wCnt1 = 0 Then
'            wCnt1 = 2
'            wCnt2 = 1
'        End If
'    End If
'    CIS3D3(0) = String(wCnt1, "▼")
'    CIS3D3(1) = String(8 - wCnt1, "▼")
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "PR"
        Call DispChange(ProcHB)
        DoEvents
        
'       ( Proc Start )
        Call Proc_Start
    
        ProcHB = "C1"
        Call DispChange(ProcHB)
        Timer3.Enabled = True
        
        With CisFrm
            Set .MSGObject = lb_msg
            Set .MSGTimer = Timer2
            .MSGInterval = 100
            .MSGStopCount = 20
            .MSGType = Left_To_Right
            .MSGText = "【 処理終了 】確認して下さい"
            
            .MSGRelate
        End With
        
        GoTo ReturnPress_Ed
    End If
    
ReturnPress_Ed:
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
    Dim H1Mode          As Boolean
    Dim H1Color         As Long
    Dim PRMode          As Boolean
    Dim PRColor         As Long
    Dim C1Mode          As Boolean
    Dim C1Color         As Long
    
    H1Mode = False:     H1Color = gPLostSel
    PRMode = False:     PRColor = gPLostSel
    C1Mode = False:     C1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "PR":      PRMode = True:  PRColor = gPGotSel
        Case "C1":      C1Mode = True:  C1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    C1_Area1.Visible = C1Mode
    
    PB_ENT.Visible = Not (PRMode Or C1Mode)
    PNL_ENT.Visible = Not (PRMode Or C1Mode)
    PB_END.Visible = Not (PRMode Or C1Mode)
    PNL_END.Visible = Not (PRMode Or C1Mode)
    
    PB_OK.Visible = C1Mode
    PNL_OK.Visible = C1Mode
    
    Dummy.Enabled = False
End Function
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
'   < 対象年月 >
    If RTrim$(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    今回対象年月を入力して下さい         "
            .MB_Button = Error
            .MBOX
            iH1_YM.SetFocus
            Exit Function
        End With
    End If
    If RTrim$(iH1_YM) < RTrim$(mZenYm) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    前回処理年月より過去の年月は不可！         "
            .MB_Button = Error
            .MBOX
            iH1_YM.SetFocus
            Exit Function
        End With
    End If
    If RTrim(mZenYm) <> "" Then
        gStr = Format(DateAdd("m", 1, CDate(Left(mZenYm, 4) & "/" & Mid(mZenYm, 5, 2) & "/01")), "yyyymm")
        If RTrim$(iH1_YM) > RTrim$(gStr) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "    " & Left(gStr, 4) & " 年 " & Mid(gStr, 5, 2) & " 月度の処理が行われていません！         "
                .MB_MSG(3) = "    　よろしいですか？      "
                .MB_Button = OK_CAN
                If Not .MBOX Then
                    iH1_YM.SetFocus
                    Exit Function
                End If
            End With
        End If
    End If
    
    With CisFun
        .MB_Lines = 4
        gStr = Left(iH1_YM, 4) & " 年 " & Mid(iH1_YM, 5, 2) & "月度 "
        If RTrim$(iH1_YM) = RTrim$(mZenYm) Then
            .MB_MSG(1) = "   " & gStr & " 【再処理】         "
        Else
            .MB_MSG(1) = "   " & gStr & " 【初回処理】         "
        End If
        .MB_MSG(3) = "    売上予想作成処理を開始します？         "
        .MB_Button = Yes_No
        If Not .MBOX Then Exit Function
    End With
    
    Head1Chk = True
End Function
'*--------------------------*
'*      処　理　開　始      *
'*--------------------------*
Private Function Proc_Start() As Boolean
    Dim wwShime    As Byte
    
    Proc_Start = False
    
    
    STime = Now()
    
    lb_msg = "【 売上予想処理中 】しばらくお待ち下さい": DoEvents
    
    On Error GoTo Proc_Start_Err
'======================================='
'   【内示】
'======================================='
    If Not NaijiProc Then Exit Function

'======================================='
'   【型費】
'======================================='
    If Not KatahiProc Then Exit Function

Proc_Start_ED:

'   < 今回内容保存 >
    Call ProcRegistry_Save
    
    On Error GoTo 0
    
    Proc_Start = True
    Exit Function

Proc_Start_Err:
    CisFun.ErrorBox "売上予想作成処理"
    End
End Function
'+===================================================+
'+     内示テーブルより売上予想作成
'+===================================================+
Private Function NaijiProc() As Boolean
    NaijiProc = False
    
    Call CisDB.DBTran(TransBegin)
    With CisDB
        .SQL = "売上予想作成"
        .StoadoCount = 7
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = OutPut用
        .ParaNo = 4:    .ParaIO = OutPut用
        .ParaNo = 5:    .ParaIO = OutPut用
        .ParaNo = 6:    .ParaIO = OutPut用
        .ParaNo = 7:    .ParaIO = OutPut用
    
        .ParaNo = 1:    .ParaValue = RTrim(iH1_YM)           ' 対象年月
        .ParaNo = 2:    .ParaValue = RTrim(iH1_Torcd)        ' 契約先
    
        .DBStored
    
        .ParaNo = 0
        If .ParaValue <> 0 Then GoTo NaijiProc_Error
    
        .ParaNo = 3: C1lb_Input(1) = Format(CisFun.Val2(C1lb_Input(1)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 4: C1lb_Output(2) = Format(CisFun.Val2(C1lb_Output(2)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 5: C1lb_Output(1) = Format(CisFun.Val2(C1lb_Output(1)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 6: C1lb_Output(3) = Format(CisFun.Val2(C1lb_Output(3)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 7:
                        If CisFun.Val2(.ParaValue) <> 0 Then
                            C1lb_Input(1) = C1lb_Input(1) & "(" & Format(CisFun.Val2(.ParaValue), "#,###") & ")"
                        End If
    End With
    Call CisDB.DBTran(TransCommit)
    
    NaijiProc = True
    Exit Function
NaijiProc_Error:
    Select Case CisDB.ParaValue
        Case 2: gStr = "売上予想ワーク　削除エラー！"
        Case 3: gStr = "売上予想未登録テーブル　削除エラー！"
        Case 4: gStr = "売上予想未決テーブル　削除エラー！"
        Case 5: gStr = "売上予想テーブル（当月）　削除エラー！"
        Case 6: gStr = "売上予想テーブル（期間外）　削除エラー！"
        Case 10: gStr = "売上予想未登録テーブル　追加エラー！"
        Case 11: gStr = "売上予想未登録テーブル　更新エラー！"
        Case 12: gStr = "売上予想ワーク　追加エラー！"
        Case 13: gStr = "売上予想ワーク　更新エラー！"
        Case 14: gStr = "売上予想未決テーブル　追加エラー！"
        Case 15: gStr = "売上予想未決テーブル　更新エラー！"
        Case 16: gStr = "売上予想テーブル　追加エラー！"
        Case 17: gStr = "売上予想テーブル（前月）　追加エラー！"
    End Select
    
    CisFun.MB_Lines = 5
    CisFun.MB_MSG(2) = "    売上予想処理　異常      "
    CisFun.MB_MSG(4) = "    ERROR NO = " & CisDB.ParaValue & " : " & RTrim(gStr)
    CisFun.MB_Button = Error
    CisFun.MBOX
    End

End Function
'+===================================================+
'+      型費テーブルより売上予想作成
'+===================================================+
Private Function KatahiProc() As Boolean
    KatahiProc = False
    
    Call CisDB.DBTran(TransBegin)
    With CisDB
        .SQL = "売上予想型費"
        .StoadoCount = 5
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = OutPut用
        .ParaNo = 4:    .ParaIO = OutPut用
        .ParaNo = 5:    .ParaIO = OutPut用
    
        .ParaNo = 1:    .ParaValue = RTrim(iH1_YM)           ' 対象年月
        .ParaNo = 2:    .ParaValue = RTrim(iH1_Torcd)        ' 契約先
    
        .DBStored
    
        .ParaNo = 0
        If .ParaValue <> 0 Then GoTo KatahiProc_Error
    
        .ParaNo = 4: C1lb_Input(2) = Format(CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 5: C1lb_Output(4) = Format(CisFun.Val2(.ParaValue), "#,###")
    End With
    Call CisDB.DBTran(TransCommit)
    
    KatahiProc = True
    Exit Function
KatahiProc_Error:
    Select Case CisDB.ParaValue
        Case 30: gStr = "売上予想テーブル（当月）　削除エラー！"
        Case 31: gStr = "売上予想テーブル　追加エラー！"
        Case 32: gStr = "売上予想テーブル（前月）　追加エラー！"
        Case 33: gStr = "売上予想テーブル　追加エラー！"
        Case 34: gStr = "売上予想テーブル　追加エラー！"
    End Select
    
    CisFun.MB_Lines = 5
    CisFun.MB_MSG(2) = "    売上予想(型費)処理　異常      "
    CisFun.MB_MSG(4) = "    ERROR NO = " & CisDB.ParaValue & " : " & RTrim(gStr)
    CisFun.MB_Button = Error
    CisFun.MBOX
    End
End Function

'*-------------------------------------------------*
'*      単価未決リスト
'*-------------------------------------------------*
Public Sub TanErr_List()
    
'   < 印刷データ読み込み >
    gSL_Select = "SELECT HN.表示品番,HT.契約先," & vbCr
    gSL_Select = gSL_Select & "TW.品番,TW.取引先,TW.受入,TW.数量1,TW.数量2,TW.数量3," & vbCr
    gSL_Select = gSL_Select & " ISNULL(TR1.略称,'') 納入先名," & vbCr
    gSL_Select = gSL_Select & " ISNULL(TR2.略称,'') 契約先名" & vbCr
    gSL_Select = gSL_Select & "  From 売上予想未決テーブル TW" & vbCr
    gSL_Select = gSL_Select & "   LEFT OUTER JOIN 品番マスタ HN" & vbCr
    gSL_Select = gSL_Select & "    ON HN.品番 = TW.品番" & vbCr
    gSL_Select = gSL_Select & "   LEFT OUTER JOIN 品番取引先マスタ HT" & vbCr
    gSL_Select = gSL_Select & "    ON HT.品番 = TW.品番" & vbCr
    gSL_Select = gSL_Select & "    And HT.取引先 = TW.取引先" & vbCr
    gSL_Select = gSL_Select & "    And HT.受入 = TW.受入" & vbCr
    gSL_Select = gSL_Select & "    And HT.品目 = 0" & vbCr
    gSL_Select = gSL_Select & "   LEFT OUTER JOIN 取引先マスタ TR1" & vbCr
    gSL_Select = gSL_Select & "    ON TR1.取引先CD = TW.取引先" & vbCr
    gSL_Select = gSL_Select & "    And TR1.取引先区分 = 0" & vbCr
    gSL_Select = gSL_Select & "   LEFT OUTER JOIN 取引先マスタ TR2" & vbCr
    gSL_Select = gSL_Select & "    ON TR2.取引先CD = HT.契約先" & vbCr
    gSL_Select = gSL_Select & "    And TR2.取引先区分 = 0" & vbCr

    gSL_Select = gSL_Select & " Order By TW.取引先,HN.表示品番,HT.契約先" & vbCr
    
    If Not UYMRead(gSL_Select) Then
       Call UYMClose
       Exit Sub
    End If
    
    Dim CisSimplePrint As New CisSimplePrint
    
    With CisSimplePrint
        .PrintTitle = "売上予想単価未決リスト"
        .Preservation = App.Title
        .PrintSize = A4
        .PrintOrientation = Portrait
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        
        .pHeadItemSet 0, 中, "納　入  先"
        .pHeadItemSet 1, 中, "受入"
        .pHeadItemSet 2, 中, "品  番"
        .pHeadItemSet 3, 右, "当月数量"
        .pHeadItemSet 4, 右, "翌月数量"
        .pHeadItemSet 5, 右, "翌々月数量"
        .pHeadItemSet 6, 中, "契約先"
    
        .pBodyItemSet 0, 20
        .pBodyItemSet 1, 4
        .pBodyItemSet 2, 25
        .pBodyItemSet 3, 8, 右
        .pBodyItemSet 4, 8, 右
        .pBodyItemSet 5, 10, 右
        .pBodyItemSet 6, 6
        
        .BodyInit
        
        If Not .CheckPrtCsv Then Exit Sub
        
        Do Until (Not UYM_RDSTS)
            
            gStr = UYM.取引先
            If Trim(UYM.納入先名) <> "" Then gStr = gStr & "：" & CisFun.Mid2(UYM.納入先名, 1, 20)
            .pBodyValueSet 0, gStr
            .pBodyValueSet 1, RTrim$(UYM.受入)
            If Trim(UYM.表示品番) = "" Then
               .pBodyValueSet 2, RTrim$(UYM.品番)
            Else
               .pBodyValueSet 2, RTrim$(UYM.表示品番)
            End If
            .pBodyValueSet 3, Format(UYM.数量1, "#,###")
            .pBodyValueSet 4, Format(UYM.数量2, "#,###")
            .pBodyValueSet 5, Format(UYM.数量3, "#,###")
            gStr = UYM.契約先
'''            If Trim(KNT.契約先名) <> "" Then gStr = gStr & "：" & CisFun.Mid2(KNT.契約先名, 1, 20)
            .pBodyValueSet 6, gStr
            
            .BodySet
            
            Call UYMReadNext
        Loop
        
        .PrintStart
    
    End With
    
End Sub
'*-------------------------------------------------*
'*      品番未登録リスト
'*-------------------------------------------------*
Public Sub HinErr_List()
    
'   < 印刷データ読み込み >
    gSL_Select = "SELECT HN.表示品番,TW.品番,TW.取引先,TW.受入," & vbCr
    gSL_Select = gSL_Select & "TW.工区,TW.管理番号,TW.内示数,TW.号補区分," & vbCr
    gSL_Select = gSL_Select & " ISNULL(TR1.略称,'') 納入先名" & vbCr
    gSL_Select = gSL_Select & "  From 売上予想未登録テーブル TW" & vbCr
    gSL_Select = gSL_Select & "   LEFT OUTER JOIN 品番マスタ HN" & vbCr
    gSL_Select = gSL_Select & "    ON HN.品番 = TW.品番" & vbCr
    gSL_Select = gSL_Select & "   LEFT OUTER JOIN 取引先マスタ TR1" & vbCr
    gSL_Select = gSL_Select & "    ON TR1.取引先CD = TW.取引先" & vbCr
    gSL_Select = gSL_Select & "    And TR1.取引先区分 = 0" & vbCr

    gSL_Select = gSL_Select & " Order By TW.取引先,TW.受入,HN.表示品番,TW.管理番号" & vbCr
    
    If Not UYHRead(gSL_Select) Then
       Call UYHClose
       Exit Sub
    End If
    
    Dim CisSimplePrint As New CisSimplePrint
    
    With CisSimplePrint
        .PrintTitle = "売上予想品番未登録リスト"
        .Preservation = App.Title
        .PrintSize = A4
        .PrintOrientation = Portrait
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        
        .pHeadItemSet 0, 中, "納　入  先"
        .pHeadItemSet 1, 中, "受入"
        .pHeadItemSet 2, 中, "品  番"
        .pHeadItemSet 3, 中, "管理番号"
        .pHeadItemSet 4, 中, "号補"
        .pHeadItemSet 5, 右, "内示数"
        .pHeadItemSet 6, 中, "工区"
    
        .pBodyItemSet 0, 20
        .pBodyItemSet 1, 4
        .pBodyItemSet 2, 25
        .pBodyItemSet 3, 8
        .pBodyItemSet 4, 4
        .pBodyItemSet 5, 10, 右
        .pBodyItemSet 6, 4
        
        .BodyInit
        
        If Not .CheckPrtCsv Then Exit Sub
        
        Do Until (Not UYH_RDSTS)
            
            gStr = UYH.取引先
            If Trim(UYH.納入先名) <> "" Then gStr = gStr & "：" & CisFun.Mid2(UYH.納入先名, 1, 20)
            .pBodyValueSet 0, gStr
            .pBodyValueSet 1, RTrim$(UYH.受入)
            If Trim(UYH.表示品番) = "" Then
               .pBodyValueSet 2, RTrim$(UYH.品番)
            Else
               .pBodyValueSet 2, RTrim$(UYH.表示品番)
            End If
            .pBodyValueSet 3, RTrim$(UYH.管理番号)
            .pBodyValueSet 4, RTrim$(UYH.号補区分)
            .pBodyValueSet 5, Format(UYH.内示数, "#,###")
            .pBodyValueSet 6, RTrim$(UYH.工区)
            
            .BodySet
            
            Call UYHReadNext
        Loop
        
        .PrintStart
    
    End With
    
End Sub

'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽
'▽                         名　称　マ　ス　タ　確　認  ( シ ス テ ム 共 通 定 数 等 )
'▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'*---------------------------------------------------------*
'*                  前　回　内　容　獲　得                  *
'*---------------------------------------------------------*
Private Sub ProcRegistry_Check()
    
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "処理内容"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "売上予想作成処理"
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)

    Call ItemsClearSYM
    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '処理内容'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値 = '" & App.EXEName & "'"
    If SYMRead(gSL_Select, 1) Then
        With SYM
            If RTrim$(.英数字1) <> "" Then
                H1lb_ZYM = Format(.英数字1, "@@@@年 @@月度")
                mZenYm = RTrim(.英数字1)
            End If
            If RTrim$(.英数字2) <> "" Then
                H1lb_ZDate = .英数字2
            End If
            If RTrim$(.英数字3) <> "" Then
                If TorNmGet(.英数字3, 2) Then
                    H1lb_KTorcd = .英数字3 & ":" & TRM_RName
                    If Trim(H1lb_KTorcd) = "" Then
                       H1lb_KTorcd = .英数字3 & ":" & TRM_Name
                    End If
                Else
                    H1lb_KTorcd = .英数字3
                End If
            End If
        End With
    End If
'自社情報
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "自社情報"
        .区分桁数 = 1
        .値 = "3"
        .制御区分 = ""
        .値名称 = "期"
        .英数字1 = "数字(1)=上期開始月"
        .数字1 = 4
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)
End Sub
'*---------------------------------------------------------*
'*                  今　回　内　容　保  存                  *
'*---------------------------------------------------------*
Private Sub ProcRegistry_Save()

    Call ItemsClearSYM
    
    With SYM
        .区分名称 = "処理内容"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "売上予想作成処理"
        .英数字1 = iH1_YM
        .英数字2 = Format(Now(), "YYYY年 MM月 DD日 ") & Format(Now(), "HH:NN")
        .英数字3 = RTrim(iH1_Torcd)
        .作成者 = gTanto
    End With

    Call SYMUpdate
End Sub
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
        .値名称 = "売上予想作成処理"
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



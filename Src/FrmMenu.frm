VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8F4FF474-3E18-11D4-AEC2-00000EA57F9E}#1.0#0"; "CISMSG~1.OCX"
Begin VB.Form FrmMenu 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "SIMPR-A System Menu"
   ClientHeight    =   10890
   ClientLeft      =   1950
   ClientTop       =   2670
   ClientWidth     =   15360
   Icon            =   "FrmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10890
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D Back_Main 
      Height          =   9945
      Left            =   135
      Top             =   765
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   17542
      BackColor       =   16761024
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
      Begin Cis3D_v60.CIS3D cmd_MainNo 
         Height          =   255
         Index           =   0
         Left            =   60
         Top             =   30
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin Cis3D_v60.CIS3D cmd_Main 
         Height          =   810
         Index           =   0
         Left            =   30
         Top             =   0
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   1429
         BackColor       =   16744576
         Caption         =   "受 信 管 理"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cBoderWidth     =   3
         cFont3D         =   1
         cAlingnment     =   7
         cButton         =   -1  'True
         cPositionY      =   60
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D7 
      Height          =   1785
      Left            =   135
      Top             =   8925
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   3149
      BackColor       =   16761024
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
      Begin Cis3D_v60.CIS3D lb_Type 
         Height          =   825
         Left            =   120
         Top             =   30
         Visible         =   0   'False
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1455
         BackColor       =   16761024
         Caption         =   "B-Type"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin Cis3D_v60.CIS3D lb_Version 
         Height          =   435
         Left            =   870
         Top             =   870
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
         BackColor       =   16761024
         Caption         =   "Ver 1.00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cBoderStyle     =   2
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   555
      Left            =   3690
      ScaleHeight     =   495
      ScaleWidth      =   11445
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   10125
      Width           =   11505
      Begin VB.Label lb_Company 
         BackStyle       =   0  '透明
         Caption         =   "XXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   105
         Width           =   4785
      End
      Begin VB.Label lb_DateTime 
         Alignment       =   1  '右揃え
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "YYYY/MM/DD hh:mm:ss"
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9090
         TabIndex        =   3
         Top             =   135
         Width           =   2295
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   9345
      Left            =   3690
      Top             =   765
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   16484
      BackColor       =   16761024
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
      cBoderColor1    =   8421504
      cBoderColor2    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D cmd_Sub 
         Height          =   705
         Index           =   2
         Left            =   7665
         Top             =   30
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   1244
         BackColor       =   16744576
         Caption         =   "XXXXXXXXX1XXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   15.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
         cButton         =   -1  'True
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D cmd_SubNo 
            Height          =   675
            Index           =   2
            Left            =   30
            Top             =   30
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   1191
            BackColor       =   12648447
            Caption         =   "Z"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   14737632
            cFont3DColor2   =   0
            cFont3D         =   2
            cAlingnment     =   7
         End
      End
      Begin Cis3D_v60.CIS3D cmd_Sub 
         Height          =   705
         Index           =   0
         Left            =   60
         Top             =   30
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   1244
         BackColor       =   12648384
         Caption         =   "XXXXXXXXX1XXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   15.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
         cButton         =   -1  'True
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D cmd_SubNo 
            Height          =   675
            Index           =   0
            Left            =   30
            Top             =   30
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   1191
            BackColor       =   12648447
            Caption         =   "Z"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   14737632
            cFont3DColor2   =   0
            cFont3D         =   2
            cAlingnment     =   7
         End
      End
      Begin Cis3D_v60.CIS3D cmd_Sub 
         Height          =   705
         Index           =   1
         Left            =   3862
         Top             =   30
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   1244
         BackColor       =   16744576
         Caption         =   "XXXXXXXXX1XXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   15.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
         cButton         =   -1  'True
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D cmd_SubNo 
            Height          =   675
            Index           =   1
            Left            =   30
            Top             =   30
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   1191
            BackColor       =   12648447
            Caption         =   "Z"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   14737632
            cFont3DColor2   =   0
            cFont3D         =   2
            cAlingnment     =   7
         End
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   0
         Left            =   30
         Top             =   750
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   10
         Left            =   5745
         Top             =   750
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   1
         Left            =   30
         Top             =   1605
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   11
         Left            =   5745
         Top             =   1605
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   2
         Left            =   30
         Top             =   2460
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   12
         Left            =   5745
         Top             =   2460
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   3
         Left            =   30
         Top             =   3315
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   13
         Left            =   5745
         Top             =   3315
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   4
         Left            =   30
         Top             =   4170
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   14
         Left            =   5745
         Top             =   4170
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   5
         Left            =   30
         Top             =   5025
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   15
         Left            =   5745
         Top             =   5025
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   6
         Left            =   30
         Top             =   5880
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   16
         Left            =   5745
         Top             =   5880
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   7
         Left            =   30
         Top             =   6735
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   17
         Left            =   5745
         Top             =   6735
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   8
         Left            =   30
         Top             =   7590
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   18
         Left            =   5745
         Top             =   7590
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   9
         Left            =   30
         Top             =   8445
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D cmd_Exec 
         Height          =   855
         Index           =   19
         Left            =   5745
         Top             =   8445
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   60
      End
   End
   Begin CisMsg_v60.CISMSG lb_System 
      Height          =   675
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   15480
      _ExtentX        =   27305
      _ExtentY        =   1191
      BackColor       =   -2147483633
      ForeColor       =   16777215
      Caption         =   "総合生産管理システム"
      BackStyle       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   26.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cAlingnment     =   7
      gGC1_Red        =   200
      gGC2_Green      =   255
      gGC3_Bule       =   150
      gGM1_Red        =   4
      gGM2_Green      =   4
      gGM3_Bule       =   2
      gGradation      =   -1  'True
   End
   Begin CisMsg_v60.CISMSG CISMSG3 
      Height          =   675
      Left            =   -30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1191
      BackColor       =   16711680
      ForeColor       =   65535
      Caption         =   "Propss"
      BackStyle       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   65535
      cBoderColor2    =   16711680
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cAlingnment     =   7
      gGM1_Red        =   4
      gGM2_Green      =   4
      gGM3_Bule       =   2
      gGradation      =   -1  'True
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      Left            =   750
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   510
      Width           =   285
   End
   Begin VB.Timer Timer1 
      Left            =   210
      Top             =   330
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   420
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   600
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Menu File 
      Caption         =   "ﾌｧｲﾙ(&F)"
      Begin VB.Menu MSql_Start 
         Caption         =   "SQLｽｹｼﾞｭｰﾙ開始"
         Visible         =   0   'False
      End
      Begin VB.Menu MSql_End 
         Caption         =   "SQLｽｹｼﾞｭｰﾙ停止"
         Visible         =   0   'False
      End
      Begin VB.Menu M_Line2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MN_ChgSet 
         Caption         =   "ﾒﾆｭｰ切替設定"
         Visible         =   0   'False
      End
      Begin VB.Menu MN_Chg 
         Caption         =   "ﾒﾆｭｰ切替(&C)"
      End
      Begin VB.Menu M_Line3 
         Caption         =   "-"
      End
      Begin VB.Menu M_END 
         Caption         =   "終了(&X)"
      End
   End
   Begin VB.Menu M_Option 
      Caption         =   "ｵﾌﾟｼｮﾝ(&O)"
      Visible         =   0   'False
      Begin VB.Menu ReaderSet 
         Caption         =   "QRﾘｰﾀﾞｰ設定"
      End
   End
   Begin VB.Menu M_Version 
      Caption         =   "ﾊﾞｰｼｮﾝ(&V)"
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  SIMPR-A ﾒﾆｭｰ
'**       フォームID    :  frmMenu
'**       処理概要      :
'**
'**       作  成  日    :   2004/01/28  By CIS
'**       変  更  日    :   2005/09/27  By CIS ﾛｸﾞｲﾝ情報入力後、環境設定の起動を行う
'**       変  更  日    :   2006/07/06  By CIS 業務ﾒﾆｭｰ切替
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ

    Dim mCIS3DObj       As CIS3D
    
    Dim mMNo            As Integer
    Dim mSNo            As Integer
    
    Dim mErCnt          As Integer
    
    Dim Menussap        As String
    Dim wMenussap       As String
    Dim wExePass        As Integer
    Dim wExec           As String
    Dim wExecKB         As String
    
    Dim mMainMax        As Integer
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
    If Not Dupli_Start Then End

    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me

' ADOﾊﾞｰｼﾞｮﾝﾁｪｯｸ
    If Not ADOCheck Then End
    lb_Company = ""
'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    CisDB.TimeOut = 7
    If Not IniGet Then
        If CisDB.ErrMsg <> "No" Then
           frmLogin.Show vbModal
           Unload frmLogin
           Set frmLogin = Nothing
           If Not gLoginSucceeded Then End
        '+------------------------------------------------------- 2005/09/27 Insert
           Shell gIniExe & "PXM0000.exe /SIMPR-A.exe"
           End
        '+------------------------------------------------------- 2005/09/27 Insert End
        Else
           End
        End If
    End If

' 名称マスタ検索　確認
    Call MeisyoChk
'
    If CisFun.Len2(CisFun.WSID) > 20 Then
        gStr = "          コンピュータ名が不正です。　" & vbCr & vbCr & _
                  "コンピュータ名の変更を行ってから、再度実行して下さい。"
        MsgBox gStr, vbCritical, ""
        End
    End If
    
''--  --  --  --  --  --  --  --  --  < 2006/07/06 ADD >
''
'   0-- メニュー情報 --0
    Call MenuInfo
    gK_EditKB = 0
''--  --  --  --  --  --  --  --  --  < 2006/07/06 ADD END >

'   #-------------------#
'   #    MENU.TXT 読込  #
'   #-------------------#
    If Not FromText Then End
    
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Dim Main_Height As Long
    
    mMainMax = gMMax
    If gMMax < 6 Then
       mMainMax = 5
    End If
    Main_Height = Back_Main.Height / (mMainMax + 1)
    If mMainMax = 9 Then
       Main_Height = 990
    End If
    cmd_Main(0).Height = Main_Height
    cmd_MainNo(0).Move cmd_Main(0).Left + 30, cmd_Main(0).Top + 30
    If mMainMax > 10 Then
       cmd_Main(0).cPositionY = 10
    End If
    For gInt = 1 To mMainMax
        Load cmd_Main(gInt)
        cmd_Main(gInt).Height = Main_Height
        cmd_Main(gInt).Move cmd_Main(gInt - 1).Left, cmd_Main(gInt - 1).Top + cmd_Main(gInt - 1).Height
        cmd_Main(gInt).Visible = True
        If mMainMax > 10 Then
           cmd_Main(gInt).cPositionY = 10
        End If
        Load cmd_MainNo(gInt)
        cmd_MainNo(gInt).Move cmd_Main(gInt - 1).Left + 30, cmd_Main(gInt - 1).Top + cmd_Main(gInt - 1).Height + 30
        cmd_MainNo(gInt).ZOrder (0)
    Next gInt


    For Each gObj In cmd_Main
        gObj.Caption = ""
    Next gObj
    For Each gObj In cmd_Sub
        gObj.Caption = ""
        gObj.Visible = False
    Next gObj
    
    For Each gObj In cmd_Exec
        gObj.Caption = ""
    Next gObj
    
    Dummy.Left = -1000
'
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    
    Call Timer1_Timer
    
    Set gPopMenu = Me.M_Option
    
    If CisDB.Server = CisFun.WSID Then
        Const VER_PLATFORM_WIN32_NT = 2&                        'WINDOWSNT,2000 XP

'       !   MajorVersion    4 : Win 9X 系(95,98,ME) Or NT4.0
        If CisFun.OSVersion.MajorVersion = 4 And _
           CisFun.OSVersion.PlatformId = VER_PLATFORM_WIN32_NT Then
            MSql_Start.Visible = True
            MSql_End.Visible = True
            M_Line2.Visible = True
        End If
    End If
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(True)
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'   #-------------------#
'   #    システム名表示
'   #-------------------#
    Call SysNameDisp
'   #-------------------#
'   #    メニュー表示    #
'   #-------------------#
    
    cmd_SubNo(0).Caption = gMS_Key(0)
    cmd_SubNo(1).Caption = gMS_Key(1)
    cmd_SubNo(2).Caption = gMS_Key(2)
    
    gMId = -1
    
    For gInt = 0 To UBound(gMM_Inf) - 1
        If gInt > mMainMax Then Exit For
        cmd_Main(gInt) = gMM_Inf(gInt)
        If Trim(gMM_Inf(gInt)) <> "" Then
           cmd_MainNo(gInt).Visible = True
           Select Case gInt
                Case Is < 10
                    gStr = Format(gInt + 1, "0")
                    cmd_MainNo(gInt).Caption = Right(gStr, 1)
                Case 10
                    cmd_MainNo(gInt).Caption = "/"
                Case 11
                    cmd_MainNo(gInt).Caption = "*"
                Case 12
                    cmd_MainNo(gInt).Caption = "-"
                Case 13
                    cmd_MainNo(gInt).Caption = "+"
                Case 14
                    cmd_MainNo(gInt).Caption = "."
           End Select
           If gMId < 0 Then
              gMId = gInt
           End If
        Else
           cmd_Main(gInt).cButton = False
        End If
    Next gInt
    
    If CisFun.Val2(Mid(Command, 1, 2)) - 1 >= 0 Then
       gMId = CisFun.Val2(Mid(Command, 1, 2)) - 1
    End If
    If gMId < 0 Then gMId = 0
    
'    If gMId = 0 Then
'       Call MenuDisp(gMId, 0)
'    Else
'       Call cmd_Main_Click(gMId)
'    End If
    Call cmd_Main_Click(gMId)
' 自社コード入力確認
    If Trim(gCompany) = "" Then
        frmUser.Show vbModal
        Unload frmUser
        Set frmUser = Nothing
        Call SysNameDisp
    End If
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                   +
'+----------------------------------------------------------+
Private Sub Form_Paint()
    Me.ScaleMode = 1
    Line1(0).X1 = 0
    Line1(0).Y1 = 0
    Line1(0).X2 = Me.ScaleWidth
    Line1(0).Y2 = 0

    Line1(1).X1 = 0
    Line1(1).Y1 = Screen.TwipsPerPixelY
    Line1(1).X2 = Me.ScaleWidth
    Line1(1).Y2 = Screen.TwipsPerPixelY
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ｕｎｌｏａd イベント》                  +
'+----------------------------------------------------------+
Private Sub Form_Unload(Cancel As Integer)
    Call cmd_Main_Click(3)
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
    
'    Static W_IVENT As String
'    W_IVENT = "Down"
    Key_Acc = False
' VbKey0
    If Key_Code = 48 Then
       Call cmd_Main_Click(9)
       Key_Acc = True
       Exit Function
    End If
' VbKey1 ～ VbKey9
    If Key_Code >= 49 And Key_Code <= (49 + 9) Then
       Call cmd_Main_Click(Key_Code - 49)
       Key_Acc = True
       Exit Function
    End If
' vbKeyDivide(/)
    If Key_Code = vbKeyDivide Then
       Call cmd_Main_Click(10)
       Key_Acc = True
       Exit Function
    End If
' vbKeyMultiply(*)
    If Key_Code = vbKeyMultiply Then
       Call cmd_Main_Click(11)
       Key_Acc = True
       Exit Function
    End If
' vbKeySubtract(-)
    If Key_Code = vbKeySubtract Then
       Call cmd_Main_Click(12)
       Key_Acc = True
       Exit Function
    End If
' vbKeyAdd(+)
    If Key_Code = vbKeyAdd Then
       Call cmd_Main_Click(13)
       Key_Acc = True
       Exit Function
    End If
' vbKeyDecimal(.)
    If Key_Code = vbKeyDecimal Then
       Call cmd_Main_Click(14)
       Key_Acc = True
       Exit Function
    End If
' VbNumKey0
    If Key_Code = 96 Then
       Call cmd_Main_Click(9)
       Key_Acc = True
       Exit Function
    End If
' VbNumKey1 ～ VbNumKey9
    If Key_Code >= 97 And Key_Code <= (97 + 9) Then
       Call cmd_Main_Click(Key_Code - 97)
       Key_Acc = True
       Exit Function
    End If
' ｻﾌﾞﾒﾆｭｰ切替
    If Shift = 0 Then                       ' < 2006/07/06 ADD >
        If gMS_Key(0) = Chr(Key_Code) Then
           Call cmd_Sub_Click(0)
           Key_Acc = True
           Exit Function
        End If
        If gMS_Key(1) = Chr(Key_Code) Then
           Call cmd_Sub_Click(1)
           Key_Acc = True
           Exit Function
        End If
        If gMS_Key(2) = Chr(Key_Code) Then
           Call cmd_Sub_Click(2)
           Key_Acc = True
           Exit Function
        End If
    End If                                  ' < 2006/07/06 ADD >
    
'    '***  矢印キー  ***
    If Key_Code = vbKeyLeft Or Key_Code = vbKeyRight Or _
       Key_Code = vbKeyUp Or Key_Code = vbKeyDown Then
        Call KeyInSelect(Key_Code)
        Key_Acc = True
        Exit Function
    End If
'
    '***  選択項目実行  ***
    If Key_Code = vbKeyReturn Then
        Call ReturnPress
        Key_Acc = True
        Exit Function
    End If
' F1 ～ F12
    If Key_Code >= 112 And Key_Code <= 123 Then
       Call FuncCall(Key_Code - 111)
       Key_Acc = True
       Exit Function
    End If

'--  --  --  --  --  --  --  --  --  < 2006/07/06 ADD >
'ﾒﾆｭｰ切替設定
    If Key_Code = vbKeyC Then
        If Shift = 7 Then
            Call MN_ChgSet_Click
            Key_Acc = True
            Exit Function
        End If
    End If
'--  --  --  --  --  --  --  --  --  < 2006/07/06 ADD END >

'
End Function
'/*---------------------------------------*/
'/*     矢印キーによる選択項目の移動
'/*---------------------------------------*/
Private Sub KeyInSelect(KeyCode As Integer)
    Dim wIndex      As Integer
    Dim wIndex2     As Integer
    Dim wEd         As Integer
    Dim wStep       As Integer
    Dim wSetFlg     As Boolean
    
' 現在有効ボタンのｲﾝﾃﾞｯｸｽを求める
    For Each mCIS3DObj In cmd_Exec
        If mCIS3DObj.cButton Then
           wIndex = mCIS3DObj.Index
           Exit For
        End If
    Next mCIS3DObj
    
    Select Case KeyCode
        Case vbKeyRight
                wIndex = wIndex + 10
                wEd = 19
                wStep = 1
        Case vbKeyLeft
                wIndex = wIndex - 10
                wEd = 0
                wStep = -1
        Case vbKeyUp
                wIndex = wIndex - 1
                wEd = 0
                wStep = -1
        Case vbKeyDown
                wIndex = wIndex + 1
                wEd = 19
                wStep = 1
    End Select
    
    If wIndex < 0 Then wIndex = 19
    If wIndex > 19 Then wIndex = 0
    
    For wIndex2 = wIndex To wEd Step wStep
        mMNo = CisFun.Val2(Trim(cmd_Exec(wIndex2).Tag))
        If Trim(gMExe(mMNo, wIndex2 + mSNo * 20)) <> "" Then
            Call Mouse_Move(wIndex2)
            wSetFlg = True
            Exit For
        End If
    Next wIndex2
    
    If Not wSetFlg Then
       For wIndex2 = 0 To 19
           mMNo = CisFun.Val2(Trim(cmd_Exec(wIndex2).Tag))
           If Trim(gMExe(mMNo, wIndex2 + mSNo * 20)) <> "" Then
              Call Mouse_Move(wIndex2)
              Exit For
           End If
       Next wIndex2
    End If
End Sub
'--  --  --  --  --  --  --  --  --  --  --  --  --  --  --  < 2006/07/06 ADD >
'*------------------------------*
'*  ( ﾒﾆｭｰﾊﾞｰ ) メニュー切替    *
'*------------------------------*
Private Sub MN_Chg_Click()
    gK_MNSelKB = 0              ' ﾒﾆｭｰ選択区分( 0:選択、1:ﾒﾆｭｰ登録 )
    Call MenuSel_Load
End Sub
'*------------------------------*
'*  ( ﾒﾆｭｰﾊﾞｰ ) メニュー切替設定 *
'*------------------------------*
Private Sub MN_ChgSet_Click()
    gK_MNSelKB = 1              ' ﾒﾆｭｰ選択区分( 0:選択、1:ﾒﾆｭｰ登録 )
    Call MenuSel_Load
End Sub
'*------------------------------*
'*  メニュー選択画面表示
'*------------------------------*
Private Function MenuSel_Load()

    FrmMenuSel.Show vbModal
    Unload FrmMenuSel
    Set FrmMenuSel = Nothing

    If gK_MNSelRtn = 1 Then
'       < メニュー切替 >
        With CisFun
            .MB_Lines = 5
            .MB_MSG(1) = "    切替を有効にするには、         "
            .MB_MSG(2) = "    　メニューを再起動して下さい         "
            .MB_Button = OK
            .MBOX
        End With
        Call M_END_Click
        Exit Function
    End If
    
    If gK_EditKB = 1 Then
'       < メニュー切替 >
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    設定内容を有効にするには、         "
            .MB_MSG(2) = "    　メニューを再起動して下さい         "
            .MB_Button = OK
            .MBOX
        End With
        Call M_END_Click
        Exit Function
    End If

End Function
'--  --  --  --  --  --  --  --  --  --  --  --  --  --  --  < 2006/07/06 ADD END >
'**************************************
'*      選択項目展開・実行(Enter)      *
'**************************************
Private Sub ReturnPress()
    For Each mCIS3DObj In cmd_Exec
        If mCIS3DObj.cButton Then
           Call Cmd_Exec_Click(mCIS3DObj.Index)
           Exit For
        End If
    Next mCIS3DObj
End Sub
'**************************************
'*      選択項目展開・実行(ﾌｧﾝｸｼｮﾝｷｰ)  *
'**************************************
Private Sub FuncCall(FKey As Integer)
    Dim FIdx        As Integer
    For Each mCIS3DObj In cmd_Exec
        mMNo = CisFun.Val2(Trim(mCIS3DObj.Tag))
        If gMExeNo(mMNo, mCIS3DObj.Index) = FKey Then
           FIdx = mCIS3DObj.Index
           Call Mouse_Move(FIdx)
           Call Cmd_Exec_Click(FIdx)
           Exit For
        End If
    Next mCIS3DObj
End Sub
'+------------------------------------------+
'+          業務　メニュー　選択
'+------------------------------------------+
Private Sub cmd_Main_Click(Index As Integer)
    
    On Error GoTo cmd_Main_click_ed:
    
    If Trim(cmd_Main(Index).Caption) = "" Then Exit Sub
    For Each gObj In cmd_Main
        If gObj.Index = Index Then
           gObj.BackColor = &HC0FFC0
        Else
           gObj.BackColor = &HFF8080
        End If
    Next gObj
    Call MenuDisp(Index, 0)
    If cmd_Main(Index).Tag = "END" Then
        Call EndProc
    End If
cmd_Main_click_ed:

End Sub
'+------------------------------------------+
'+          業務　メニュー　選択(No)
'+------------------------------------------+
Private Sub cmd_MainNo_Click(Index As Integer)
    Call cmd_Main_Click(Index)
End Sub
'+------------------------------------------+
'+          サブ　メニュー　選択
'+------------------------------------------+
Private Sub cmd_Sub_Click(Index As Integer)
    If Index > 2 Then Exit Sub
    If Trim(cmd_Sub(Index).Caption) = "" Or Not cmd_Sub(Index).Visible Then Exit Sub
    Call MenuDisp(CisFun.Val2(cmd_Sub(Index).Tag), Index)
End Sub
Private Sub cmd_SubNo_Click(Index As Integer)
    Call cmd_Sub_Click(Index)
End Sub

Private Sub EndProc()
    End
End Sub
'+----------------------------------------------------------+
'+      Timer1(現在時刻表示)  《 Timer イベント》
'+----------------------------------------------------------+
Private Sub Timer1_Timer()
    lb_DateTime = Format(Now(), "yyyy/mm/dd ") & _
                  Format(Now(), "hh:mm:ss")
    
End Sub
'+----------------------------------------------------------+
'+      cmd_Exec(実行ﾓｼﾞｭｰﾙボタン)
'+----------------------------------------------------------+
'《 MouseMove イベント》
Private Sub cmd_Exec_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Mouse_Move(Index)
End Sub
'+----------------------------------------------------------+
'+      各選択項目のMouse Move 処理
'+----------------------------------------------------------+
Private Sub Mouse_Move(Index As Integer)
    On Error GoTo Mouse_Move_Err
    
    mMNo = CisFun.Val2(Trim(cmd_Exec(Index).Tag))
    gLong = Index + mSNo * 20
    If Trim(gMExe(mMNo, gLong)) = "" Then Exit Sub
    
    ' 前回選択項目クリア
    For Each mCIS3DObj In cmd_Exec
        If mCIS3DObj.Index = Index Then
           If Not mCIS3DObj.cButton Then
              mCIS3DObj.cButton = True
              mCIS3DObj.cFont3D = 0
              mCIS3DObj.ForeColor = &HFF&
           End If
        Else
           If mCIS3DObj.cButton Then
              mCIS3DObj.cButton = False
              mCIS3DObj.cBoderStyle = 2
              mCIS3DObj.cFont3D = 1
              mCIS3DObj.ForeColor = &H80000008
           End If
        End If
    Next mCIS3DObj
'
Mouse_Move_Next:
    On Error GoTo 0
    
    Exit Sub
Mouse_Move_Err:
    GoTo Mouse_Move_Next
End Sub
'《 Click イベント》
'*** プログラム起動 ***
Private Sub Cmd_Exec_Click(Index As Integer)
    Dim wSNo        As Long
    ' ﾒﾆｭｰ番号　退避
    gInt = CisFun.Val2(cmd_Exec(Index).Tag)
    wSNo = Index + mSNo * 20
    If Trim(cmd_Exec(Index)) = "" Or gMExe(gInt, wSNo) = "" Then Exit Sub

    On Error GoTo PGMStart_Err
    wExec = gMExe(gInt, wSNo)
    wExecKB = Trim(gMExeKB(gInt, wSNo))
    On Error Resume Next
    If InStr(1, wExec, "EDIG00.exe") <> 0 Then
        ChDir gEdiExeDir
    End If
    On Error GoTo PGMStart_Err
    
    wExePass = 0
    CisDB.PassLink
    
    Select Case wExecKB
        Case "", "S": Call Shell(gIniExe & wExec, 1)
        Case "X":
                With CisFun
                     .ShellApl = gIniExe & wExec
                     .ShellTask Freeze
                End With
        Case "P":      Call Shell(wExec, 1)
        Case "E":
                                gLong = ShellExecute(Me.hWnd, _
                                vbNullString, _
                                wExec, _
                                vbNullString, _
                                Left$(CisFun.WinRoot, 3), _
                                SW_SHOWNORMAL)
        Case "L":       Call LogDisp(wExec)
        Case "D"
            Call Shell(gIniExe & wExec & " " & gInt + 1, 1)
            Call M_END_Click
    End Select
    
    If wExecKB = "X" Then
       Call IniGet
       Call SysNameDisp
    End If
    On Error GoTo 0
    Exit Sub
PGMStart_Err:
    If Err = 53 Then
        With CisFun
            .MB_Lines = 7
            .MB_MSG(2) = "     指定された実行ﾌｧｲﾙが存在しません       "
            If Trim(gMExeKB(gInt, Index)) = "" Then
               .MB_MSG(5) = "    実行ファイル名 ＝ " & StrConv(gIniExe & wExec, vbWide) & "         "
            Else
               .MB_MSG(5) = "    実行ファイル名 ＝ " & StrConv(wExec, vbWide) & "         "
            End If
            .MB_Title = "プログラム起動エラー"
            .MB_Button = Error
            Call .MBOX
        End With
        
        Resume Next
    End If
    Call CisFun.ErrorBox
    Resume Next
End Sub

'+-----------------------------------------+
'+-----------------------------------------+
'+          メニュー内容表示                +
'+-----------------------------------------+
'+-----------------------------------------+
Private Sub MenuDisp(MNo As Integer, SNo As Integer)
    Dim wSeKFlg     As Boolean
    Dim wMNo        As Byte
    Dim wSCnt       As Long     ' ｻﾌﾞﾒﾆｭｰ件数
    Dim wCNo        As Long
    
    For Each mCIS3DObj In cmd_Exec
        If Trim(mCIS3DObj.Caption) <> "" Then
            mCIS3DObj.Caption = ""
            mCIS3DObj.Tag = ""
            mCIS3DObj.cButton = False
            mCIS3DObj.cBoderStyle = 2
            mCIS3DObj.cFont3D = 1
            mCIS3DObj.ForeColor = &H80000008
        End If
    Next mCIS3DObj

    Erase gMExeNo
    
    For Each gObj In cmd_Sub
        gObj.Caption = ""
        gObj.Tag = ""
        gObj.Visible = False
    Next gObj
    wSCnt = 0
    For gInt = 0 To UBound(gMS_Inf, 2) - 1
        If Trim(gMS_Inf(MNo, gInt)) <> "" Then
           wSCnt = wSCnt + 1
        End If
    Next gInt
    If wSCnt > 1 Then
        For gInt = 0 To UBound(gMS_Inf, 2) - 1
            If Trim(gMS_Inf(MNo, gInt)) <> "" Then
               cmd_Sub(gInt).Caption = Trim(gMS_Inf(MNo, gInt))
               cmd_Sub(gInt).Tag = MNo
               cmd_Sub(gInt).Visible = True
            End If
        Next gInt
        For Each gObj In cmd_Sub
            If gObj.Index = SNo Then
               gObj.BackColor = &HC0FFC0
            Else
               gObj.BackColor = &HFF8080
            End If
        Next gObj
    End If
    
    mSNo = SNo      ' ｻﾌﾞﾒﾆｭｰ番号退避
    
    wCNo = SNo * 20
    
    For gInt = 0 To 19
        If gInt > 19 Then Exit For
        cmd_Exec(gInt).Tag = MNo
        If Trim(gMTitle(MNo, wCNo)) <> "" Then
           If Trim(gMExe(MNo, wCNo)) <> "" Then
              wMNo = wMNo + 1
              gMExeNo(MNo, gInt) = wMNo
              gStr = Format(wMNo, "##") & "  "
              If wMNo < 13 Then
                 cmd_Exec(gInt).Caption = "F" & Mid(gStr, 1, 2) & "." & gMTitle(MNo, wCNo)
              Else
                 cmd_Exec(gInt).Caption = " " & Mid(gStr, 1, 2) & "." & gMTitle(MNo, wCNo)
              End If
              If wMNo = 1 Then
                 cmd_Exec(gInt).cButton = True
                 cmd_Exec(gInt).cFont3D = 0
                 cmd_Exec(gInt).ForeColor = &HFF&
              End If
           Else
              cmd_Exec(gInt).Caption = gMTitle(MNo, wCNo)
           End If
        End If
        wCNo = wCNo + 1
    Next gInt

End Sub
'+----------------------------------+
'+      ログ内容表示
'+----------------------------------+
Private Sub LogDisp(FName As String)
    Dim wFName()     As String
    Dim strFName     As String
    Dim FIdx         As Long
    Dim wMaxFile     As String
    
    Erase wFName
    ReDim wFName(1)
'
    strFName = Dir(gIniExe & "Log\" & FName & "*.*")       ' 最初のファイル名を返します。
    If Trim(strFName) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "       表示するログファイルがありません。          "
            .MB_Title = "ﾛｸﾞ表示"
            Call .MBOX
        End With
       Exit Sub
    End If
    
    Do While strFName <> ""                ' ループを開始します。
        FIdx = FIdx + 1
        ReDim Preserve wFName(FIdx)
        wFName(FIdx - 1) = strFName

        strFName = Dir                     ' 次のファイル名を返します。
    Loop
'' 受信ﾌｧｲﾙ名分類(降順:最新ﾌｧｲﾙのみ)
    For FIdx = 1 To UBound(wFName) - 1
        If wFName(0) < wFName(FIdx) Then
           wFName(0) = wFName(FIdx)
        End If
    Next FIdx

    wMaxFile = gIniExe & "Log\" & wFName(0)
    
    gLong = ShellExecute(Me.hWnd, vbNullString, "NOTEPAD.EXE", wMaxFile, _
                                Left(CisFun.WinRoot, 3), _
                                SW_SHOWNORMAL)
    
End Sub
'#-------------------#
'#    システム名表示
'#-------------------#
Private Sub SysNameDisp()
    lb_System.Caption = Trim(gSysName)
    lb_Company.Caption = gCompName
End Sub
'/*---------------------------------------*/
'/*     ﾌｧｲﾙ － 終了
'/*---------------------------------------*/
Private Sub M_END_Click()
    Call CisDB.DBDISConnect
    Call CisFrm.UnLoadDisp
    End
End Sub
'/*---------------------------------------*/
'/*     ｵﾌﾟｼｮﾝ
'/*---------------------------------------*/
'/*---------------------------------------*/
'/*     バージョン情報
'/*---------------------------------------*/
Private Sub M_Version_Click()
    SysMenuAbout.Show vbModal
End Sub
'/*---------------------------------------*/
'/*     ﾌｧｲﾙ － SQLｽｹｼﾞｭｰﾙ停止
'/*---------------------------------------*/
Private Sub MSql_Start_Click()
    Call CisFrm.MousePT(11)
    
    With CisFun
        If .TackCheck("SQLAgent.exe") = 0 Then
            .MB_Lines = 4
            .MB_MSG(2) = "       ＳＱＬスケジュール起動されています。          "
            .MB_Title = "Sql-Agent"
            .MBOX
        
            Call CisFrm.MousePT(1)
            Exit Sub
        End If
    End With
    DoEvents
    
    Call Shell(gIniExe & "PSQAStart.exe 3", 1)

'   * 時間遅延 *
    Static sw_Date      As Date
    sw_Date = DateAdd("s", 12, Now())
    Do Until (sw_Date <= Now())
    Loop
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "       ＳＱＬスケジュールの起動が開始されました。          "
        .MB_Title = "Sql-Agent"
        .MBOX
    
        Call CisFrm.MousePT(1)
    End With

End Sub
'/*---------------------------------------*/
'/*     ﾌｧｲﾙ － SQLｽｹｼﾞｭｰﾙ停止
'/*---------------------------------------*/
Private Sub MSql_End_Click()
    DoEvents
    Call CisFrm.MousePT(11)
    
    With CisFun
        gInt = .TackCheck("SQLAgent.exe", 終了する)
        gInt = .TackCheck("WINOA386.MOD", 終了する)
            
        .MB_Lines = 4
            
        Select Case gInt
            Case 0:     .MB_MSG(2) = "       ＳＱＬスケジュールが正常終了しました。          "
            Case 1:     .MB_MSG(2) = "       ＳＱＬスケジュールは起動されていません。          "
            Case Else:  .MB_MSG(2) = "       ＳＱＬスケジュールが異常終了しました。          "
        End Select
        .MB_Title = "Sql-Agent"
        .MBOX
    End With
    
    Call CisFrm.MousePT(1)
End Sub
'#-----------------------#
'#   ADOﾊﾞｰｼﾞｮﾝﾁｪｯｸ      #
'#-----------------------#
Private Function ADOCheck() As Boolean
    Static ADO_Version          As String

    ADOCheck = False
    ADO_Version = ""
  
    With CisFun
        
'       ( ADO Version Check )
        .Reg_RootKey = HKEY_LOCAL_MACHINE
        .Reg_MainKey = "Software\"
        .Reg_RegKey = "Microsoft\"
        .Reg_SubKey = "DataAccess\"
        .Reg_ValueName = "VersionX"
    
        .Reg_ValueData = ""
    
        .RegProcess REG_DATA_GET
    
        ADO_Version = .Reg_ValueData
    
        If .Val2(Mid(ADO_Version, 1, 3)) >= 2.5 Then
            ADOCheck = True
            Exit Function
        End If
    
'       ( ADO FullInstallVersion Check )
        .Reg_ValueName = "FullInstallVer"
        .Reg_ValueData = ""
    
        .RegProcess REG_DATA_GET
    
        ADO_Version = .Reg_ValueData
    
        If .Val2(Mid(ADO_Version, 1, 3)) >= 2.5 Then
            ADOCheck = True
            Exit Function
        End If
    
    End With

' ADOｲﾝｽﾄｰﾙ確認ﾒｯｾｰｼﾞ表示
    gStr = "                    データベース接続エラー        　" & vbCr & vbCr & _
           "    CD-ROM内の 'mdac_typ.exe'を実行後、再度実行して下さい。"
    MsgBox gStr, vbCritical, "SIMPR-A"
End Function
'#-----------------------#
'#   名称マスタ存在確認  #
'#-----------------------#
Private Sub MeisyoChk()
    
    gSL_Select = "SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = '名称検索'"
    gSL_Select = gSL_Select & "   AND 制御区分  = 'X'"
    gSL_Select = gSL_Select & "   AND 値   = '01'"
    If Not SYMRead(gSL_Select, 1) Then
        Call ItemsClearSYM
        With SYM
            .区分名称 = "名称検索"
            .区分桁数 = 2
            .値 = "01"
            .値名称 = "名称検索"
            .英数字1 = "区   分   検   索"
            .英数字2 = "値"
            .英数字3 = "区分名"
            .制御区分 = "X"
        End With
        If Not SYMReadChk Then
            Call SYMInsert
        End If
    End If
    '+---------------------+
    '+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
    '+---------------------+
    With CisFun
        .INI_Section = "名称マスタ"
        .INI_FileID = gIniExe & "Combo.ini"
        .INI_Default = ""
        .INI_Key = "TYPE"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_Key = "TYPE"
           .INI_String = "SQL"
           .PutIni
           .INI_Key = "CNT"
           .INI_String = "3"
           .PutIni
           .INI_Key = "DATA1"
           .INI_String = "SELECT 値,値名称 FROM 名称マスタ"
           .PutIni
           .INI_Key = "DATA2"
           .INI_String = "WHERE 区分名称 = '?'   AND 区分タイプ = ''"
           .PutIni
           .INI_Key = "DATA3"
           .INI_String = "ORDER BY 値"
           .PutIni
        End If
    End With
    
End Sub
'--  < 2006/07/06 ADD >  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --
'
'*------------------*
'*   メニュー情報   *
'*------------------*
Private Function MenuInfo()
    MenuInfo = False

    With CisFun
        .INI_FileID = gIniExe & gIniName
        .INI_Section = .WSID
        .INI_Default = ""
        
'   < メニュー数 >
        .INI_Key = "MenuKinds"
        .GetIni
        gKind_Cnt = .Val2(.INI_String)
        If gKind_Cnt = 0 Then gKind_Cnt = 1
    
'   < 最終表示メニュー >
        .INI_Key = "MenuLast"
        .GetIni
        gK_Last = .Val2(.INI_String)
    
    End With
    
    MN_Chg.Visible = Not (gKind_Cnt = 1)
    M_Line3.Visible = Not (gKind_Cnt = 1)

    MenuInfo = True
End Function
'
'--  < 2006/07/06 ADD END >  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --

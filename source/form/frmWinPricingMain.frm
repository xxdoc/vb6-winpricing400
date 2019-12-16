VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWinPricingMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmWinPricingMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "ใบโอนสุกรเข้าเรือนขาย"
      Height          =   495
      Left            =   3480
      TabIndex        =   52
      Top             =   6300
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ใบเปลี่ยนสถานะสุกร"
      Height          =   495
      Left            =   3480
      TabIndex        =   51
      Top             =   5790
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ใบโอนสุกร"
      Height          =   495
      Left            =   3480
      TabIndex        =   50
      Top             =   5280
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ใบโอนวัตถุดิบ"
      Height          =   495
      Left            =   3480
      TabIndex        =   49
      Top             =   4770
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ใบเบิกวัตถุดิบ"
      Height          =   495
      Left            =   3480
      TabIndex        =   48
      Top             =   4260
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ใบสุกรคลอด"
      Height          =   495
      Left            =   3480
      TabIndex        =   47
      Top             =   3750
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ใบนำเข้าวัตถุดิบ"
      Height          =   495
      Left            =   3480
      TabIndex        =   46
      Top             =   3240
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command2 
      Caption         =   "เคลียร์ Interim"
      Height          =   495
      Left            =   3480
      TabIndex        =   45
      Top             =   2730
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "สร้าง Interim"
      Height          =   495
      Left            =   3480
      TabIndex        =   44
      Top             =   2220
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3300
      Top             =   1110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":24B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":2D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":3666
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":3980
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":425A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":4B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":540E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":5CE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   795
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1402
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand SSCommand1 
         Height          =   555
         Left            =   9660
         TabIndex        =   38
         Top             =   6390
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   979
         _Version        =   131073
         PictureFrames   =   1
         Picture         =   "frmWinPricingMain.frx":5FF7
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin VB.Label lblDateTime 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   315
         Left            =   9390
         TabIndex        =   37
         Top             =   30
         Width           =   2505
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7755
      Left            =   0
      TabIndex        =   0
      Top             =   780
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   13679
      _Version        =   131073
      BackStyle       =   1
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   240
         ScaleHeight     =   1215
         ScaleWidth      =   1185
         TabIndex        =   53
         Top             =   4320
         Width           =   1185
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin MSComctlLib.TreeView trvMain 
         Height          =   3045
         Left            =   240
         TabIndex        =   1
         Top             =   1230
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   5371
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "JasmineUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdPasswd 
         Height          =   465
         Left            =   330
         TabIndex        =   40
         Top             =   7170
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   820
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   465
         Left            =   1920
         TabIndex        =   39
         Top             =   7170
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   820
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblVersion 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   36
         Top             =   6600
         Width           =   3045
      End
      Begin VB.Label lblUserGroup 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   35
         Top             =   6090
         Width           =   3045
      End
      Begin VB.Label lblUserName 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   34
         Top             =   5580
         Width           =   3045
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   795
      Left            =   3450
      TabIndex        =   3
      Top             =   810
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   1402
      _Version        =   131073
      BackStyle       =   1
   End
   Begin Threed.SSFrame fraMain 
      Height          =   4875
      Left            =   6000
      TabIndex        =   4
      Top             =   7860
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8599
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdMainEmployee 
         Height          =   765
         Left            =   900
         TabIndex        =   9
         Top             =   2820
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMainReport 
         Height          =   765
         Left            =   900
         TabIndex        =   8
         Top             =   3600
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMainEnterprise 
         Height          =   765
         Left            =   900
         TabIndex        =   7
         Top             =   480
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMainCustomer 
         Height          =   765
         Left            =   900
         TabIndex        =   6
         Top             =   1260
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMainSupplier 
         Height          =   765
         Left            =   900
         TabIndex        =   5
         Top             =   2040
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraAdmin 
      Height          =   3615
      Left            =   9960
      TabIndex        =   10
      Top             =   7680
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6376
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdAdminReport 
         Height          =   765
         Left            =   900
         TabIndex        =   13
         Top             =   2190
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdUser 
         Height          =   765
         Left            =   900
         TabIndex        =   12
         Top             =   1410
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdUserGroup 
         Height          =   765
         Left            =   900
         TabIndex        =   11
         Top             =   630
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraMaster 
      Height          =   5505
      Left            =   3690
      TabIndex        =   14
      Top             =   7410
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9710
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdMaster6 
         Height          =   765
         Left            =   900
         TabIndex        =   69
         Top             =   4290
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPricePlanMaster 
         Height          =   765
         Left            =   900
         TabIndex        =   63
         Top             =   1170
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMaster3 
         Height          =   765
         Left            =   900
         TabIndex        =   19
         Top             =   3510
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMaster2 
         Height          =   765
         Left            =   900
         TabIndex        =   18
         Top             =   1950
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMaster1 
         Height          =   765
         Left            =   900
         TabIndex        =   17
         Top             =   390
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMaster5 
         Height          =   765
         Left            =   900
         TabIndex        =   16
         Top             =   2730
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMaster4 
         Height          =   765
         Left            =   900
         TabIndex        =   15
         Top             =   2730
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraInventory 
      Height          =   5625
      Left            =   9120
      TabIndex        =   20
      Top             =   7440
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9922
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdAdjust 
         Height          =   765
         Left            =   900
         TabIndex        =   42
         Top             =   3600
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdTransfer 
         Height          =   765
         Left            =   900
         TabIndex        =   25
         Top             =   2820
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdInventoryReport 
         Height          =   765
         Left            =   900
         TabIndex        =   24
         Top             =   4380
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdRawMatterial 
         Height          =   765
         Left            =   900
         TabIndex        =   23
         Top             =   480
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdImport 
         Height          =   765
         Left            =   900
         TabIndex        =   22
         Top             =   1260
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExport 
         Height          =   765
         Left            =   900
         TabIndex        =   21
         Top             =   2040
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraPig 
      Height          =   4935
      Left            =   8700
      TabIndex        =   26
      Top             =   7560
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8705
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdPigAdjustment 
         Height          =   765
         Left            =   900
         TabIndex        =   43
         Top             =   2850
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPigWeek 
         Height          =   765
         Left            =   900
         TabIndex        =   41
         Top             =   510
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPigBirth 
         Height          =   765
         Left            =   900
         TabIndex        =   29
         Top             =   1290
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPigTransfer 
         Height          =   765
         Left            =   900
         TabIndex        =   28
         Top             =   2070
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPigReport 
         Height          =   765
         Left            =   900
         TabIndex        =   27
         Top             =   3630
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraLedger 
      Height          =   4455
      Left            =   5430
      TabIndex        =   30
      Top             =   2100
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7858
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdCurrencyExchange 
         Height          =   765
         Left            =   900
         TabIndex        =   75
         Top             =   750
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdLedgerReport 
         Height          =   765
         Left            =   900
         TabIndex        =   33
         Top             =   3090
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdBuy 
         Height          =   765
         Left            =   900
         TabIndex        =   32
         Top             =   2310
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSell 
         Height          =   765
         Left            =   900
         TabIndex        =   31
         Top             =   1530
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraGoldStock 
      Height          =   4365
      Left            =   5460
      TabIndex        =   54
      Top             =   7500
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7699
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdGoldWage 
         Height          =   765
         Left            =   900
         TabIndex        =   58
         Top             =   1410
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdGldDailyPrice 
         Height          =   765
         Left            =   900
         TabIndex        =   57
         Top             =   630
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdGldSaleBuy 
         Height          =   765
         Left            =   900
         TabIndex        =   56
         Top             =   2190
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdGldReport 
         Height          =   765
         Left            =   900
         TabIndex        =   55
         Top             =   2970
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraPackage 
      Height          =   3615
      Left            =   11610
      TabIndex        =   59
      Top             =   7320
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6376
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdPackageReport 
         Height          =   765
         Left            =   900
         TabIndex        =   62
         Top             =   2220
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFeature 
         Height          =   765
         Left            =   900
         TabIndex        =   61
         Top             =   660
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSoc 
         Height          =   765
         Left            =   900
         TabIndex        =   60
         Top             =   1440
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraPerson 
      Height          =   4095
      Left            =   5520
      TabIndex        =   64
      Top             =   7770
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7223
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdSalarySlipt 
         Height          =   765
         Left            =   900
         TabIndex        =   68
         Top             =   2070
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMoneyPerson 
         Height          =   765
         Left            =   900
         TabIndex        =   67
         Top             =   1290
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDataPerson 
         Height          =   765
         Left            =   900
         TabIndex        =   66
         Top             =   510
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdReportPerson 
         Height          =   765
         Left            =   900
         TabIndex        =   65
         Top             =   2850
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraProduction 
      Height          =   4365
      Left            =   5040
      TabIndex        =   70
      Top             =   -3120
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7699
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdProductionReport 
         Height          =   765
         Left            =   900
         TabIndex        =   74
         Top             =   2970
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdProductionEstimate 
         Height          =   765
         Left            =   900
         TabIndex        =   73
         Top             =   2190
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdProductionFormula 
         Height          =   765
         Left            =   900
         TabIndex        =   72
         Top             =   630
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdProductionJob 
         Height          =   765
         Left            =   900
         TabIndex        =   71
         Top             =   1410
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmWinPricingMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"

Private m_Sp As CSystemParam
Private MustAsk As Boolean
Private m_HasActivate As Boolean
Private m_Rs  As ADODB.Recordset
Private m_TableName As String

Public HeaderText As String
Private m_XCollection As CXCollection
Private m_MustAsk As Boolean

Private Sub InitMainTreeview()
Dim Node As Node
Dim NewNodeID As String
   trvMain.Nodes.Clear
   trvMain.Font.Name = GLB_FONT
   trvMain.Font.Size = 14
   trvMain.Font.Bold = False

   Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE, MapText("ระบบงานทั้งหมด"), 1)
   Node.Expanded = True
   Node.Selected = True
   
   '==
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-0", MapText("ระบบข้อมูลผู้ใช้งาน"), 4, 4)
   Node.Expanded = False
   '==
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-1", MapText("ระบบข้อมูลหลัก"), 2, 2)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-2", MapText("ระบบข้อมูลส่วนกลาง"), 6, 6)
   Node.Expanded = False

   If glbGuiConfigs.VerifyGuiConfig("HR_VIEW") Then
      Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-8", MapText("ระบบบริหารฝ่ายบุคคล"), 12, 12)
      Node.Expanded = False
   End If

   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-7", MapText("ระบบแพคเกจสินค้า/บริการ"), 5, 5)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-3", MapText("ระบบบริหารคลัง"), 3, 3)
   Node.Expanded = False

   If glbGuiConfigs.VerifyGuiConfig("PRODUCTION_VIEW") Then
      Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-9", MapText("ระบบการผลิต"), 10, 10)
      Node.Expanded = False
   End If
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-5", MapText("ระบบบริหารบัญชี"), 8, 8)
   Node.Expanded = False
End Sub

Private Sub InitFormLayout()
   Call InitNormalLabel(lblUserName, MapText("ผู้ใช้ : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblUserGroup, MapText("กลุ่มผู้ใช้ : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblVersion, MapText("เวอร์ชัน : ") & glbParameterObj.Version & " (" & glbParameterObj.ProgramOwner & ") ", RGB(0, 0, 255))
   Call InitNormalLabel(lblDateTime, "", RGB(0, 0, 255))
   lblDateTime.BackStyle = 1
   lblDateTime.BackColor = RGB(255, 255, 255)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPasswd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   cmdUserGroup.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdUser.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdAdminReport.Picture = LoadPicture(glbParameterObj.MainButton)
      
   cmdMaster1.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMaster2.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMaster3.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMaster4.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMaster5.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdPricePlanMaster.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMaster6.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdMainEnterprise.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMainCustomer.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMainSupplier.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMainReport.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMainEmployee.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdRawMatterial.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdImport.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdExport.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdTransfer.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdAdjust.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdInventoryReport.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdPigWeek.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdPigBirth.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdPigTransfer.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdPigAdjustment.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdPigReport.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdCurrencyExchange.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdBuy.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdSell.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdLedgerReport.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdGldDailyPrice.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdGldSaleBuy.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdGldReport.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdGoldWage.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdFeature.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdSoc.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdPackageReport.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdDataPerson.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMoneyPerson.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdSalarySlipt.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdReportPerson.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdProductionFormula.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdProductionJob.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdProductionEstimate.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdProductionReport.Picture = LoadPicture(glbParameterObj.MainButton)

   Me.Caption = glbGuiConfigs.ShowWindowCaption(glbParameterObj.ProgramOwner)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   Call InitMainButton(cmdUserGroup, MapText("ข้อมูลกลุ่มผู้ใช้งาน"))
   Call InitMainButton(cmdUser, MapText("ข้อมูลผู้ใช้งาน"))
   Call InitMainButton(cmdAdminReport, MapText("รายงานข้อมูลผู้ใช้งาน"))

   Call InitMainButton(cmdMaster1, MapText("ข้อมูลหลักส่วนกลาง"))
   Call InitMainButton(cmdMaster2, MapText("ข้อมูลหลักระบบคลัง"))
   Call InitMainButton(cmdMaster3, MapText("ข้อมูลหลักระบบฝ่ายบุคคล"))
   Call InitMainButton(cmdMaster4, MapText("ข้อมูลหลักระบบบริหารบัญชี"))
   Call InitMainButton(cmdMaster5, MapText("ข้อมูลหลักงานขาย"))
   cmdMaster5.Visible = False
   Call InitMainButton(cmdPricePlanMaster, MapText("ข้อมูลหลักแพคเกจสินค้า/บริการ"))
   Call InitMainButton(cmdMaster6, MapText("ข้อมูลหลักระบบการผลิต"))

   Call InitMainButton(cmdMainEnterprise, MapText("ข้อมูลบริษัท"))
   Call InitMainButton(cmdMainCustomer, MapText("ข้อมูลลูกค้า"))
   Call InitMainButton(cmdMainSupplier, MapText("ข้อมูลซัพพลายเออร์"))
   Call InitMainButton(cmdMainEmployee, MapText("ข้อมูลพนักงาน"))
   Call InitMainButton(cmdMainReport, MapText("รายงานข้อมูลกลาง"))
   
   Call InitMainButton(cmdRawMatterial, MapText("ข้อมูลสินค้าและวัตถุดิบ"))
   Call InitMainButton(cmdImport, MapText("ข้อมูลการรับเข้าวัตถุดิบ"))
   Call InitMainButton(cmdExport, MapText("ข้อมูลการเบิกวัตถุดิบ"))
   Call InitMainButton(cmdTransfer, MapText("ข้อมูลการโอนย้ายวัตถุดิบ"))
   Call InitMainButton(cmdAdjust, MapText("ข้อมูลการปรับยอดคลัง"))
   Call InitMainButton(cmdInventoryReport, MapText("รายงานระบบคลัง"))
   
   Call InitMainButton(cmdPigWeek, MapText("ข้อมูลรหัสสัปดาห์เกิดสุกร"))
   Call InitMainButton(cmdPigBirth, MapText("ข้อมูลสุกรคลอด"))
   Call InitMainButton(cmdPigTransfer, MapText("ข้อมูลการโอนย้ายสุกร"))
   Call InitMainButton(cmdPigAdjustment, MapText("ข้อมูลการปรับยอดสุกร"))
   Call InitMainButton(cmdPigReport, MapText("รายงานระบบบริหารสุกร"))
   
   Call InitMainButton(cmdCurrencyExchange, MapText("ข้อมูลอัตราการแลกเปลี่ยนเงินตรา"))
   Call InitMainButton(cmdBuy, MapText("ระบบงานซื้อ (รายจ่าย)"))
   Call InitMainButton(cmdSell, MapText("ระบบงานขาย"))
   Call InitMainButton(cmdLedgerReport, MapText("รายงานระบบบัญชี"))
   
   Call InitMainButton(cmdGldDailyPrice, MapText("ราคาทองประจำวัน"))
   Call InitMainButton(cmdGoldWage, MapText("ข้อมูลค่าแรงช่างทอง"))
   Call InitMainButton(cmdGldSaleBuy, MapText("ระบบซื้อขายทอง"))
   Call InitMainButton(cmdGldReport, MapText("รายงานระบบร้านทอง"))
   
   Call InitMainButton(cmdFeature, MapText("ข้อมูลสินค้า/บริการ"))
   Call InitMainButton(cmdSoc, MapText("ข้อมูลแพคเกจสินค้า/บริการ"))
   Call InitMainButton(cmdPackageReport, MapText("รายงานแพคเกจสินค้า/บริการ"))
   
   Call InitMainButton(cmdDataPerson, MapText("ข้อมูลพนักงาน"))
   Call InitMainButton(cmdMoneyPerson, MapText("เงินยืมส่วนบุคคล"))
   Call InitMainButton(cmdSalarySlipt, MapText("สลิปเงินเดือน"))
   Call InitMainButton(cmdReportPerson, MapText("รายงานระบบฝ่ายบุคคล"))
   
   Call InitMainButton(cmdProductionFormula, MapText("ข้อมูลสูตรการผลิต"))
   Call InitMainButton(cmdProductionJob, MapText("ข้อมูลใบสั่งผลิต"))
   Call InitMainButton(cmdProductionEstimate, MapText("ข้อมูลใบประเมินราคา"))
   Call InitMainButton(cmdProductionReport, MapText("รายงานระบบการผลิต"))

   Call InitMainButton(cmdExit, MapText("ออก"))
   Call InitMainButton(cmdPasswd, MapText("โปรแกรม"))
   
   Picture1.Visible = glbGuiConfigs.VerifyGuiConfig("LOGO_VIEW")
   If glbGuiConfigs.VerifyGuiConfig("LOGO_VIEW") Then
      Picture1.Picture = LoadPicture(glbParameterObj.CompanyLogo)
   End If
   
   Call InitMainTreeview
End Sub

Private Sub cmdAdjust_Click()
   If Not VerifyAccessRight("INVENTORY_ADJUST") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmInventoryDoc4
   frmInventoryDoc4.Show 1

   Unload frmInventoryDoc4
   Set frmInventoryDoc4 = Nothing
End Sub

Private Sub cmdAdminReport_Click()
   If Not VerifyAccessRight("ADMIN_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmSummaryReport.HeaderText = cmdAdminReport.Caption
   frmSummaryReport.MasterMode = 1
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub

Private Sub cmdBuy_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ใบเสนอราคาซื้อ", "-", "ใบส่งของ", "-", "ใบกำกับภาษี", "-", "ใบเสร็จรับเงิน", "-", "ใบเพิ่มหนี้", "-", "ใบลดหนี้", "-", "ใบบรรจุหีบห่อ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not glbGuiConfigs.VerifyGuiConfig("BUY_QUOATATION_EXEC", True) Then
         Exit Sub
      End If
      frmBillingDoc1.DocumentType = 13
   ElseIf lMenuChosen = 3 Then
      If Not glbGuiConfigs.VerifyGuiConfig("BUY_DO_EXEC", True) Then
         Exit Sub
      End If
      frmBillingDoc1.DocumentType = 7
   ElseIf lMenuChosen = 5 Then
      If Not glbGuiConfigs.VerifyGuiConfig("BUY_INVOICE_EXEC", True) Then
         Exit Sub
      End If
      frmBillingDoc1.DocumentType = 11
   ElseIf lMenuChosen = 7 Then
      If Not glbGuiConfigs.VerifyGuiConfig("BUY_RECEIPT_EXEC", True) Then
         Exit Sub
      End If
      frmBillingDoc1.DocumentType = 8
   ElseIf lMenuChosen = 9 Then
      If Not glbGuiConfigs.VerifyGuiConfig("BUY_DBN_EXEC", True) Then
         Exit Sub
      End If
      frmBillingDoc1.DocumentType = 10
   ElseIf lMenuChosen = 11 Then
      If Not glbGuiConfigs.VerifyGuiConfig("BUY_CDN_EXEC", True) Then
         Exit Sub
      End If
      frmBillingDoc1.DocumentType = 9
       
    ElseIf lMenuChosen = 13 Then
      If Not glbGuiConfigs.VerifyGuiConfig("BUY_PKGLST_EXEC", True) Then
         Exit Sub
      End If
      frmBillingDoc1.DocumentType = 15
   
   End If
   
   
   
   frmBillingDoc1.Area = 2
   Load frmBillingDoc1
   frmBillingDoc1.Show 1

   Unload frmBillingDoc1
   Set frmBillingDoc1 = Nothing
End Sub

Private Sub cmdCurrencyExchange_Click()
'   If Not VerifyAccessRight("ACCOUNT_EXCHANGE") Then
'      Call EnableForm(Me, True)
'      Exit Sub
'   End If

   Load frmCurrency
   frmCurrency.Show 1
   
   Unload frmCurrency
   Set frmCurrency = Nothing

End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdExport_Click()
   If Not VerifyAccessRight("INVENTORY_EXPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmInventoryDoc2
   frmInventoryDoc2.Show 1
   
   Unload frmInventoryDoc2
   Set frmInventoryDoc2 = Nothing
End Sub

Private Sub cmdFeature_Click()
   If Not VerifyAccessRight("PRICEPLAN_FEATURE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmFeature
   frmFeature.Show 1
   
   Unload frmFeature
   Set frmFeature = Nothing
End Sub

Private Sub cmdGldDailyPrice_Click()
   If Not VerifyAccessRight("GOLD_PRICE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmGoldDailyPrice
   frmGoldDailyPrice.Show 1
   
   Unload frmGoldDailyPrice
   Set frmGoldDailyPrice = Nothing
End Sub

Private Sub cmdGldSaleBuy_Click()
   If Not VerifyAccessRight("GOLD_SELLBUY") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmGoldSellBuy
   frmGoldSellBuy.Show 1

   Unload frmGoldSellBuy
   Set frmGoldSellBuy = Nothing
End Sub

Private Sub cmdGoldWage_Click()
   If Not VerifyAccessRight("GOLD_WAGE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmGoldWage
   frmGoldWage.Show 1
   
   Unload frmGoldWage
   Set frmGoldWage = Nothing
End Sub

Private Sub cmdImport_Click()
   If Not VerifyAccessRight("INVENTORY_IMPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmInventoryDoc1
   frmInventoryDoc1.Show 1
   
   Unload frmInventoryDoc1
   Set frmInventoryDoc1 = Nothing
End Sub

Private Sub cmdInventoryReport_Click()
   If Not VerifyAccessRight("INVENTORY_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmSummaryReport.HeaderText = cmdInventoryReport.Caption
   frmSummaryReport.MasterMode = 4
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub

Private Sub cmdLedgerReport_Click()
'   If Not VerifyAccessRight("LEDGER_REPORT") Then
'      Call EnableForm(Me, True)
'      Exit Sub
'   End If

   frmSummaryReport.HeaderText = cmdLedgerReport.Caption
   frmSummaryReport.MasterMode = 5
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub

Private Sub cmdMainCustomer_Click()
   If Not VerifyAccessRight("MAIN_CUSTOMER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmCustomer
   frmCustomer.Show 1
   
   Unload frmCustomer
   Set frmCustomer = Nothing
End Sub

Private Sub cmdMainEmployee_Click()
   If Not VerifyAccessRight("MAIN_EMPLOYEE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If


   Load frmEmployee
   frmEmployee.Show 1
   
   Unload frmEmployee
   Set frmEmployee = Nothing
End Sub

Private Sub cmdMainEnterprise_Click()
   If Not VerifyAccessRight("MAIN_ENTERPRISE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmAddEditEnterprise.ShowMode = SHOW_ADD
   frmAddEditEnterprise.HeaderText = cmdMainEnterprise.Caption
   Load frmAddEditEnterprise
   frmAddEditEnterprise.Show 1
   
   Unload frmAddEditEnterprise
   Set frmAddEditEnterprise = Nothing
End Sub

Private Sub cmdMainReport_Click()
   If Not VerifyAccessRight("MAIN_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmSummaryReport.HeaderText = cmdMainReport.Caption
   frmSummaryReport.MasterMode = 3
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub

Private Sub cmdMainSupplier_Click()
   If Not VerifyAccessRight("MAIN_SUPPLIER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmSupplier
   frmSupplier.Show 1
   
   Unload frmSupplier
   Set frmSupplier = Nothing
End Sub

Private Sub cmdMaster1_Click()
   If Not VerifyAccessRight("MASTER_MAIN") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmMasterMain.HeaderText = cmdMaster1.Caption
   frmMasterMain.MasterMode = 3
   Load frmMasterMain
   frmMasterMain.Show 1
   
   Unload frmMasterMain
   Set frmMasterMain = Nothing
End Sub

Private Sub cmdMaster2_Click()
   If Not VerifyAccessRight("MASTER_INVENTORY") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmMasterMain.HeaderText = cmdMaster2.Caption
   frmMasterMain.MasterMode = 1
   Load frmMasterMain
   frmMasterMain.Show 1
   
   Unload frmMasterMain
   Set frmMasterMain = Nothing
End Sub

Private Sub cmdMaster3_Click()
   If Not VerifyAccessRight("MASTER_PERSON") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmMasterMain.HeaderText = cmdMaster3.Caption
   frmMasterMain.MasterMode = 2
   Load frmMasterMain
   frmMasterMain.Show 1
   
   Unload frmMasterMain
   Set frmMasterMain = Nothing
End Sub

Private Sub cmdMaster4_Click()
'   If glbParameterObj.ProgramOwner = XEROX_OWNER Then
      If Not VerifyAccessRight("MASTER_LEDGER") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If

      frmMasterMain.HeaderText = cmdMaster4.Caption
      frmMasterMain.MasterMode = 7
      Load frmMasterMain
      frmMasterMain.Show 1

      Unload frmMasterMain
      Set frmMasterMain = Nothing
'   ElseIf glbParameterObj.ProgramOwner = PLAZA_OWNER Then
'      If Not VerifyAccessRight("MASTER_LEDGER") Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
'
'      frmMasterMain.HeaderText = cmdMaster4.Caption
'      frmMasterMain.MasterMode = 7
'      Load frmMasterMain
'      frmMasterMain.Show 1
'
'      Unload frmMasterMain
'      Set frmMasterMain = Nothing
'   End If
End Sub

Private Sub cmdMaster6_Click()
   If Not VerifyAccessRight("MASTER_PRODUCTION") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmMasterMain.HeaderText = cmdMaster6.Caption
   frmMasterMain.MasterMode = 8
   Load frmMasterMain
   frmMasterMain.Show 1

   Unload frmMasterMain
   Set frmMasterMain = Nothing
End Sub

Private Sub cmdPackageReport_Click()
   If Not VerifyAccessRight("PRICEPLAN_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmSummaryReport.HeaderText = cmdPackageReport.Caption
   frmSummaryReport.MasterMode = 2
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub

Private Sub cmdPasswd_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("เปลี่ยนรหัสผ่าน", "-", "ปรับราคาเฉลี่ย", "-", "ไมเกรตข้อมูลเก่า")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      Load frmChangePassword
      frmChangePassword.Show 1
      
      Unload frmChangePassword
      Set frmChangePassword = Nothing
   ElseIf lMenuChosen = 3 Then
      Load frmReArrangeDoc
      frmReArrangeDoc.Show 1
      
      Unload frmReArrangeDoc
      Set frmReArrangeDoc = Nothing
   ElseIf lMenuChosen = 5 Then
      Load frmMigrate
      frmMigrate.Show 1
      
      Unload frmMigrate
      Set frmMigrate = Nothing
   End If
End Sub

Private Sub cmdPigReport_Click()
   frmSummaryReport.HeaderText = cmdPigReport.Caption
   frmSummaryReport.MasterMode = 5
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub

Private Sub cmdPricePlanMaster_Click()
   If Not VerifyAccessRight("MASTER_PACKAGE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmMasterMain.HeaderText = cmdPricePlanMaster.Caption
   frmMasterMain.MasterMode = 6
   Load frmMasterMain
   frmMasterMain.Show 1
   
   Unload frmMasterMain
   Set frmMasterMain = Nothing
End Sub

Private Sub cmdProductionEstimate_Click()
   If Not VerifyAccessRight("PRODUCT_ESTIMATE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmJob.JobDocType = 2
   Load frmJob
   frmJob.Show 1
   
   Unload frmJob
   Set frmJob = Nothing
End Sub

Private Sub cmdProductionFormula_Click()
   If Not VerifyAccessRight("PRODUCT_FORMULA") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmFormula
   frmFormula.Show 1
   
   Unload frmFormula
   Set frmFormula = Nothing
End Sub

Private Sub cmdProductionJob_Click()
   If Not VerifyAccessRight("PRODUCT_JOB") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmJob.JobDocType = 1
   Load frmJob
   frmJob.Show 1
   
   Unload frmJob
   Set frmJob = Nothing
End Sub

Private Sub cmdRawMatterial_Click()
   If Not VerifyAccessRight("INVENTORY_PART") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmPartItem
   frmPartItem.Show 1
   
   Unload frmPartItem
   Set frmPartItem = Nothing
End Sub

Private Sub cmdSell_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ใบเสนอราคาขาย", "-", "ใบรับงาน/สั่งงาน (PO)", "-", "ใบส่งของ", "-", "ใบกำกับภาษี", "-", "ใบเสร็จรับเงิน", "-", "ใบเพิ่มหนี้", "-", "ใบลดหนี้", "-", "ใบสรุปวางบิล")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      If Not glbGuiConfigs.VerifyGuiConfig("SELL_QUOATATION_EXEC", True) Then
         Exit Sub
      End If
      
      frmBillingDoc1.DocumentType = 14
   ElseIf lMenuChosen = 3 Then
      If Not glbGuiConfigs.VerifyGuiConfig("SELL_PO_EXEC", True) Then
         Exit Sub
      End If
      
      frmBillingDoc1.DocumentType = 12
   ElseIf lMenuChosen = 5 Then
      If Not glbGuiConfigs.VerifyGuiConfig("SELL_DO_EXEC", True) Then
         Exit Sub
      End If
      
      frmBillingDoc1.DocumentType = 1
   ElseIf lMenuChosen = 7 Then
      If Not glbGuiConfigs.VerifyGuiConfig("SELL_INVOICE_EXEC", True) Then
         Exit Sub
      End If
      
      frmBillingDoc1.DocumentType = 5
   ElseIf lMenuChosen = 9 Then
      If Not glbGuiConfigs.VerifyGuiConfig("SELL_RECEIPT_EXEC", True) Then
         Exit Sub
      End If
      
      frmBillingDoc1.DocumentType = 2
   ElseIf lMenuChosen = 11 Then
      If Not glbGuiConfigs.VerifyGuiConfig("SELL_DBN_EXEC", True) Then
         Exit Sub
      End If
      
      frmBillingDoc1.DocumentType = 4
   ElseIf lMenuChosen = 13 Then
      If Not glbGuiConfigs.VerifyGuiConfig("SELL_CDN_EXEC", True) Then
         Exit Sub
      End If
      
      frmBillingDoc1.DocumentType = 3
   ElseIf lMenuChosen = 15 Then
      If Not glbGuiConfigs.VerifyGuiConfig("SELL_BILLS_EXEC", True) Then
         Exit Sub
      End If
      
      frmBillingDoc1.DocumentType = 6
   End If
   
   frmBillingDoc1.Area = 1
   Load frmBillingDoc1
   frmBillingDoc1.Show 1

   Unload frmBillingDoc1
   Set frmBillingDoc1 = Nothing
End Sub

Private Sub cmdSoc_Click()
   If Not VerifyAccessRight("PRICEPLAN_PACKAGE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmSoc
   frmSoc.Show 1
   
   Unload frmSoc
   Set frmSoc = Nothing
End Sub

Private Sub cmdTransfer_Click()
   If Not VerifyAccessRight("INVENTORY_TRANSFER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmInventoryDoc3
   frmInventoryDoc3.Show 1

   Unload frmInventoryDoc3
   Set frmInventoryDoc3 = Nothing
End Sub

Private Sub cmdUser_Click()
   If Not VerifyAccessRight("ADMIN_USER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmUser
   frmUser.Show 1
   
   Unload frmUser
   Set frmUser = Nothing
End Sub

Private Sub cmdUserGroup_Click()

   If Not VerifyAccessRight("ADMIN_GROUP") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
      
   Load frmUserGroup
   frmUserGroup.Show 1
   
   Unload frmUserGroup
   Set frmUserGroup = Nothing
End Sub

Private Sub Command1_Click()
   Call glbDatabaseMngr.ConnectLegacyDatabase("C:\COST\COST_MTP.MDB", "", "", glbErrorLog)
   Call glbDaily.ImportToInterim(True)
   glbErrorLog.LocalErrorMsg = "Import successfully"
   glbErrorLog.ShowUserError
   Call glbDatabaseMngr.DisConnectLegacyDatabase
End Sub

Private Sub Command2_Click()
   Call glbDaily.ClearInterim(True)
   glbErrorLog.LocalErrorMsg = "Clear successfully"
   glbErrorLog.ShowUserError
End Sub

Private Sub Command3_Click()
Dim D As CLegacy_h
Dim iCount As Long
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim Dt As CLegacy_d
Dim Ivd As CInventoryDoc

   Call EnableForm(Me, False)
   
   Set D = New CLegacy_h
   D.DOCUMENT_ID = 2
   Call glbDaily.QueryLegacyH(D, m_Rs, iCount, IsOK, glbErrorLog)
   Set D = Nothing
   
   Set TempRs = New ADODB.Recordset
   While Not m_Rs.EOF
      Set D = New CLegacy_h
      D.LEGACY_H_ID = NVLI(m_Rs("LEGACY_H_ID"), -1)
      D.QueryFlag = 1
      Call glbDaily.QueryLegacyH(D, TempRs, iCount, IsOK, glbErrorLog)
      Call D.PopulateFromRS(1, TempRs)
      
      Set Ivd = New CInventoryDoc
      Call glbDaily.CreateInventoryDoc2(D, Ivd, "Y")
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, True, glbErrorLog)
      Set Ivd = Nothing
      
      m_Rs.MoveNext
   Wend
   
   Set TempRs = Nothing
   Set D = Nothing
   Call EnableForm(Me, True)
End Sub

Private Sub Command4_Click()
Dim D As CLegacy_h
Dim iCount As Long
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim Dt As CLegacy_d
Dim Ivd As CInventoryDoc

   Call EnableForm(Me, False)
   
   Set D = New CLegacy_h
   D.DOCUMENT_ID = 4
   Call glbDaily.QueryLegacyH(D, m_Rs, iCount, IsOK, glbErrorLog)
   Set D = Nothing
   
   Set TempRs = New ADODB.Recordset
   While Not m_Rs.EOF
      Set D = New CLegacy_h
      D.LEGACY_H_ID = NVLI(m_Rs("LEGACY_H_ID"), -1)
      D.QueryFlag = 1
      Call glbDaily.QueryLegacyH(D, TempRs, iCount, IsOK, glbErrorLog)
      Call D.PopulateFromRS(1, TempRs)
      
      Set Ivd = New CInventoryDoc
      Call glbDaily.CreateInventoryDoc4(D, Ivd, "Y")
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, True, glbErrorLog)
      Set Ivd = Nothing
      
      m_Rs.MoveNext
   Wend
   
   Set TempRs = Nothing
   Set D = Nothing
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
Dim OKClick As Boolean
Dim iCount As Long

   If Not m_HasActivate Then
      m_HasActivate = True
      Call PatchDB
      
      Load frmLogin
      frmLogin.Show 1
      
      OKClick = frmLogin.OKClick
      
      Unload frmLogin
      Set frmLogin = Nothing
      
      glbEnterPrise.ENTERPRISE_ID = -1
      Call glbEnterPrise.QueryData(m_Rs, iCount)
      If Not m_Rs.EOF Then
         Call glbEnterPrise.PopulateFromRS(1, m_Rs)
      End If
      
      If Not OKClick Then
         m_MustAsk = False
         Unload Me
      Else
         Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
      End If
   End If
End Sub

Private Sub Form_Load()
   m_MustAsk = True
   Call InitFormLayout
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If m_MustAsk Then
      glbErrorLog.LocalErrorMsg = MapText("ท่านต้องการออกจากโปรแกรมใช่หรือไม่")
      If glbErrorLog.AskMessage = vbYes Then
         Cancel = False
      Else
         Cancel = True
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call ReleaseAll
   Set m_Rs = Nothing
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False

   lblDateTime.Caption = "                                                    "
   lblDateTime.Caption = DateToStringExtEx3(Now)
   lblUserName.Caption = MapText("ผู้ใช้ : ") & " " & glbUser.USER_NAME
   lblUserGroup.Caption = MapText("กลุ่มผู้ใช้ : ") & " " & glbUser.GROUP_NAME
      
  ' Timer1.Enabled = True
End Sub

Private Sub trvMain_NodeClick(ByVal Node As MSComctlLib.Node)
   If Node Is Nothing Then
      Exit Sub
   End If
   
   fraAdmin.Visible = False
   fraMaster.Visible = False
   fraMain.Visible = False
   fraInventory.Visible = False
   fraPig.Visible = False
   fraLedger.Visible = False
   fraGoldStock.Visible = False
   fraPackage.Visible = False
   fraProduction.Visible = False
   fraPerson.Visible = False
   
   pnlHeader.Caption = Node.Text
   If Node.Key = ROOT_TREE & " 1-0" Then
        fraAdmin.Left = 4710
        fraAdmin.Top = 2190
        fraAdmin.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-1" Then
        fraMaster.Left = 4710
        fraMaster.Top = 2190
        fraMaster.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-2" Then
        fraMain.Left = 4710
        fraMain.Top = 2190
        fraMain.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-3" Then
        fraInventory.Left = 4710
        fraInventory.Top = 2190
        fraInventory.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-4" Then
        fraPig.Left = 4710
        fraPig.Top = 2190
        fraPig.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-5" Then
        fraLedger.Left = 4710
        fraLedger.Top = 2190
        fraLedger.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-6" Then
        fraGoldStock.Left = 4710
        fraGoldStock.Top = 2190
        fraGoldStock.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-7" Then
        fraPackage.Left = 4710
        fraPackage.Top = 2190
        fraPackage.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-8" Then
        fraPerson.Left = 4710
        fraPerson.Top = 2190
        fraPerson.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-9" Then
        fraProduction.Left = 4710
        fraProduction.Top = 2190
        fraProduction.Visible = True
   End If
End Sub
Private Sub cmdDataPerson_Click()
   If Not VerifyAccessRight("HR_PERSON") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmDataPerson
   frmDataPerson.Show 1
   
   Unload frmDataPerson
   Set frmDataPerson = Nothing
End Sub

Private Sub cmdMoneyPerson_Click()
   If Not VerifyAccessRight("HR_MONEY") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmMoneyPerson
   frmMoneyPerson.Show 1
   
   Unload frmMoneyPerson
   Set frmMoneyPerson = Nothing
End Sub

Private Sub cmdSalarySlipt_Click()
   If Not VerifyAccessRight("HR_SLIPTSALARY") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmSliptSalary
   frmSliptSalary.Show 1
   
   Unload frmSliptSalary
   Set frmSliptSalary = Nothing
End Sub

Private Sub cmdReportPerson_Click()
   If Not VerifyAccessRight("HR_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmSummaryReport.HeaderText = cmdReportPerson.Caption
   frmSummaryReport.MasterMode = 6
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub


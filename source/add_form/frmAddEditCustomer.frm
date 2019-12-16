VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditCustomer 
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15765
   Icon            =   "frmAddEditCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   15765
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8640
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   21975
      _ExtentX        =   38761
      _ExtentY        =   15240
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSFrame SSFrame3 
         Height          =   1455
         Left            =   12000
         TabIndex        =   71
         Top             =   2280
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2566
         _Version        =   131073
         Caption         =   "SSFrame3"
         Begin Threed.SSOption ssoRound 
            Height          =   375
            Left            =   240
            TabIndex        =   73
            Top             =   840
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "ssoRound"
         End
         Begin Threed.SSOption ssoVolume 
            Height          =   375
            Left            =   240
            TabIndex        =   72
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "ssoVolume"
         End
      End
      Begin VB.ComboBox cboRateType 
         Height          =   315
         Left            =   8940
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   3840
         Width           =   4485
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   3735
         Left            =   15000
         TabIndex        =   55
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   6588
         _Version        =   131073
         Caption         =   "SSFrame2"
         Begin prjFarmManagement.uctlTextBox txtConBag1 
            Height          =   435
            Left            =   1560
            TabIndex        =   12
            Top             =   600
            Width           =   795
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtConBag2 
            Height          =   435
            Left            =   1560
            TabIndex        =   13
            Top             =   1080
            Width           =   795
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtConBag3 
            Height          =   435
            Left            =   1560
            TabIndex        =   14
            Top             =   1560
            Width           =   795
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtConKg1 
            Height          =   435
            Left            =   2640
            TabIndex        =   18
            Top             =   600
            Width           =   795
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtConKg2 
            Height          =   435
            Left            =   2640
            TabIndex        =   19
            Top             =   1080
            Width           =   795
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtConKg3 
            Height          =   435
            Left            =   2640
            TabIndex        =   20
            Top             =   1560
            Width           =   795
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtConBag4 
            Height          =   435
            Left            =   2160
            TabIndex        =   15
            Top             =   2040
            Width           =   555
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtConKg4 
            Height          =   435
            Left            =   2880
            TabIndex        =   21
            Top             =   2040
            Width           =   555
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtConBag5 
            Height          =   435
            Left            =   2160
            TabIndex        =   16
            Top             =   2520
            Width           =   555
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtConKg5 
            Height          =   435
            Left            =   2880
            TabIndex        =   22
            Top             =   2520
            Width           =   555
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtConBag6 
            Height          =   435
            Left            =   2160
            TabIndex        =   17
            Top             =   3000
            Width           =   555
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtConKg6 
            Height          =   435
            Left            =   2880
            TabIndex        =   23
            Top             =   3000
            Width           =   555
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtPRO_OTHER1_NAME 
            Height          =   435
            Left            =   600
            TabIndex        =   75
            Top             =   2040
            Width           =   1555
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtPRO_OTHER2_NAME 
            Height          =   435
            Left            =   600
            TabIndex        =   79
            Top             =   2520
            Width           =   1560
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtPRO_OTHER3_NAME 
            Height          =   435
            Left            =   600
            TabIndex        =   80
            Top             =   3000
            Width           =   1560
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin VB.Label lblCon4 
            Caption         =   "Label1"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label lblCon5 
            Caption         =   "Label1"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label lblCon6 
            Caption         =   "Label1"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label lblKg 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            Height          =   255
            Left            =   2640
            TabIndex        =   67
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblBath3 
            Caption         =   "Label1"
            Height          =   255
            Left            =   3480
            TabIndex        =   66
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblBath4 
            Caption         =   "Label1"
            Height          =   255
            Left            =   3480
            TabIndex        =   65
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblBath5 
            Caption         =   "Label1"
            Height          =   255
            Left            =   3480
            TabIndex        =   64
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label lblBath6 
            Caption         =   "Label1"
            Height          =   255
            Left            =   3480
            TabIndex        =   63
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lblBath2 
            Caption         =   "Label1"
            Height          =   255
            Left            =   3480
            TabIndex        =   62
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblBath1 
            Caption         =   "Label1"
            Height          =   255
            Left            =   3480
            TabIndex        =   61
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblCon3 
            Caption         =   "Label1"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblCon2 
            Caption         =   "Label1"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblCon1 
            Caption         =   "Label1"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblBag 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            Height          =   255
            Left            =   1200
            TabIndex        =   56
            Top             =   240
            Width           =   1155
         End
      End
      Begin prjFarmManagement.uctlTextBox txtMaxCredit 
         Height          =   435
         Left            =   10320
         TabIndex        =   46
         Top             =   1920
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlSaleByLookup 
         Height          =   465
         Left            =   1860
         TabIndex        =   11
         Top             =   3720
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   820
      End
      Begin VB.ComboBox cboEnterpriseType 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6990
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2820
         Width           =   3495
      End
      Begin VB.ComboBox cboBusinessType 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2820
         Width           =   3495
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   24
         Top             =   5040
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1470
         Width           =   4635
         _ExtentX        =   12356
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtShortName 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   1575
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtEmail 
         Height          =   435
         Left            =   1860
         TabIndex        =   6
         Top             =   1920
         Width           =   6885
         _ExtentX        =   12356
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWebSite 
         Height          =   435
         Left            =   1860
         TabIndex        =   7
         Top             =   2370
         Width           =   6945
         _ExtentX        =   16960
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBusinessDesc 
         Height          =   450
         Left            =   1860
         TabIndex        =   10
         Top             =   3270
         Width           =   9225
         _ExtentX        =   16907
         _ExtentY        =   794
      End
      Begin prjFarmManagement.uctlTextBox txtCredit 
         Height          =   435
         Left            =   4980
         TabIndex        =   2
         Top             =   1020
         Width           =   555
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   120
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin prjFarmManagement.uctlTextBox txtDiscountPercent 
         Height          =   435
         Left            =   6630
         TabIndex        =   3
         Top             =   1020
         Width           =   675
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2115
         Left            =   150
         TabIndex        =   25
         Top             =   5610
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   3731
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditCustomer.frx":27A2
         Column(2)       =   "frmAddEditCustomer.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditCustomer.frx":290E
         FormatStyle(2)  =   "frmAddEditCustomer.frx":2A6A
         FormatStyle(3)  =   "frmAddEditCustomer.frx":2B1A
         FormatStyle(4)  =   "frmAddEditCustomer.frx":2BCE
         FormatStyle(5)  =   "frmAddEditCustomer.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditCustomer.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   19245
         _ExtentX        =   33946
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtExpCode 
         Height          =   435
         Left            =   8070
         TabIndex        =   5
         Top             =   1470
         Width           =   1515
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCreditLimit 
         Height          =   435
         Left            =   8580
         TabIndex        =   44
         Top             =   1020
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeekCreditLimit 
         Height          =   435
         Left            =   10320
         TabIndex        =   52
         Top             =   2370
         Width           =   1515
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   465
         Left            =   1860
         TabIndex        =   82
         Top             =   4200
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   820
      End
      Begin Threed.SSCheck chkFreePriceFlag 
         Height          =   435
         Left            =   10200
         TabIndex        =   85
         Top             =   1440
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "chkFreePriceFlag"
      End
      Begin Threed.SSCheck chkCalPriceDlcCenterFlag 
         Height          =   435
         Left            =   12000
         TabIndex        =   84
         Top             =   1800
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "chkCalPriceDlcCenterFlag"
      End
      Begin Threed.SSCheck chkCalPricePartCenterFlag 
         Height          =   435
         Left            =   12000
         TabIndex        =   83
         Top             =   1440
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "chkCalPricePartCenterFlag"
      End
      Begin VB.Label lblLocationLookup 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   81
         Top             =   4200
         Width           =   1605
      End
      Begin VB.Label lblRateType 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7440
         TabIndex        =   70
         Top             =   3840
         Width           =   1365
      End
      Begin Threed.SSCommand cmdEditCon 
         Height          =   525
         Left            =   13560
         TabIndex        =   68
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomer.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkCheckCashFlag 
         Height          =   435
         Left            =   13560
         TabIndex        =   54
         Top             =   1080
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblWeekCreditLimit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8880
         TabIndex        =   53
         Top             =   2340
         Width           =   1395
      End
      Begin Threed.SSCheck chkCheckCreditFlag 
         Height          =   435
         Left            =   12000
         TabIndex        =   51
         Top             =   1080
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdEditCredit 
         Height          =   525
         Left            =   5640
         TabIndex        =   50
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomer.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkSuspendSales 
         Height          =   435
         Left            =   10200
         TabIndex        =   49
         Top             =   1080
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   495
         Left            =   11280
         TabIndex        =   48
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblMaxCredit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   9000
         TabIndex        =   47
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblCreditLimit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7320
         TabIndex        =   45
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblExpCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6810
         TabIndex        =   43
         Top             =   1590
         Width           =   1155
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   3450
         TabIndex        =   1
         Top             =   1020
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomer.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblResponseBy 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         TabIndex        =   42
         Top             =   3780
         Width           =   1695
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   29
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomer.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   30
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   27
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   26
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomer.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   28
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomer.frx":3EB8
         ButtonStyle     =   3
      End
      Begin VB.Label lblDiscountPercent 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5460
         TabIndex        =   40
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3990
         TabIndex        =   39
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblBusinessDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   38
         Top             =   3390
         Width           =   1695
      End
      Begin VB.Label lblWebsite 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   37
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   36
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblShortName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         TabIndex        =   35
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblEnterpriseType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5400
         TabIndex        =   34
         Top             =   2880
         Width           =   1485
      End
      Begin VB.Label lblBusinessType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   33
         Top             =   2880
         Width           =   1485
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   32
         Top             =   1560
         Width           =   1575
      End
   End
   Begin prjFarmManagement.uctlTextBox uctlTextBox3 
      Height          =   435
      Left            =   1440
      TabIndex        =   59
      Top             =   0
      Width           =   795
      _ExtentX        =   4471
      _ExtentY        =   767
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddEditCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Customer As CCustomer
Private m_Employees As Collection
Private m_Customers As Collection
Private m_Locations As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Private FileName As String
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_Customer.CUSTOMER_ID = ID
      If Not glbDaily.QueryCustomer(m_Customer, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Customer.PopulateFromRS(4, m_Rs)
      
      txtEmail.Text = m_Customer.EMAIL
      txtWebSite.Text = m_Customer.WEBSITE
      cboBusinessType.ListIndex = IDToListIndex(cboBusinessType, m_Customer.CUSTOMER_TYPE)
      cboEnterpriseType.ListIndex = IDToListIndex(cboEnterpriseType, m_Customer.CUSTOMER_GRADE)
      txtShortName.Text = m_Customer.CUSTOMER_CODE
      txtBusinessDesc.Text = m_Customer.BUSINESS_DESC
      txtCredit.Text = m_Customer.Credit
      txtDiscountPercent.Text = m_Customer.NORMAL_DISCOUNT
      uctlSaleByLookup.MyCombo.ListIndex = IDToListIndex(uctlSaleByLookup.MyCombo, m_Customer.RESPONSE_BY)
      txtExpCode.Text = m_Customer.EXP_CODE
      txtCreditLimit.Text = m_Customer.CREDIT_LIMIT
      txtMaxCredit.Text = m_Customer.MAX_CREDIT
      chkSuspendSales.Value = FlagToCheck(m_Customer.SUSPEND_SALES)
      chkCheckCreditFlag.Value = FlagToCheck(m_Customer.CHECK_CREDIT_FLAG)
      txtWeekCreditLimit.Text = m_Customer.WEEK_CREDIT_LIMIT
      chkCheckCashFlag.Value = FlagToCheck(m_Customer.CASH_FLAG)
      chkFreePriceFlag.Value = FlagToCheck(m_Customer.FREE_PRICE_FLAG)
      chkCalPricePartCenterFlag.Value = FlagToCheck(m_Customer.CAL_PRICE_PART_CENTER_FLAG)
      chkCalPriceDlcCenterFlag.Value = FlagToCheck(m_Customer.CAL_PRICE_DLC_CENTER_FLAG)
      uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, m_Customer.LOCATION_ID)
      
      txtConBag1.Text = m_Customer.PRO_COMMISSION_BAG
      txtConBag2.Text = m_Customer.PRO_CHEER_BAG
      txtConBag3.Text = m_Customer.PRO_DST_BAG
      txtConBag4.Text = m_Customer.PRO_OTHER1_BAG
      txtConBag5.Text = m_Customer.PRO_OTHER2_BAG
      txtConBag6.Text = m_Customer.PRO_OTHER3_BAG
      txtConKg1.Text = m_Customer.PRO_COMMISSION_KG
      txtConKg2.Text = m_Customer.PRO_CHEER_KG
      txtConKg3.Text = m_Customer.PRO_DST_KG
      txtConKg4.Text = m_Customer.PRO_OTHER1_KG
      txtConKg5.Text = m_Customer.PRO_OTHER2_KG
      txtConKg6.Text = m_Customer.PRO_OTHER3_KG
      
      txtPRO_OTHER1_NAME.Text = m_Customer.PRO_OTHER1_NAME
      txtPRO_OTHER2_NAME.Text = m_Customer.PRO_OTHER2_NAME
      txtPRO_OTHER3_NAME.Text = m_Customer.PRO_OTHER3_NAME
   
      cboRateType.ListIndex = IDToListIndex(cboRateType, m_Customer.PRICE_THINK_TYPE)
      Call GetValue(m_Customer.CAL_RATE_DELIVERY_TYPE)
      
      Dim NAME As CName
      Dim CstName As CCustomerName
      If (Not m_Customer.CstNames Is Nothing) And (m_Customer.CstNames.Count > 0) Then
         Set CstName = m_Customer.CstNames(1)
         Set NAME = CstName.NAME
         txtName.Text = NAME.LONG_NAME
      Else
         txtName.Text = ""
      End If
   Else
      ShowMode = SHOW_ADD
   End If
   
   If ShowMode = SHOW_ADD Then
      Dim Acc As CAccount
      Dim Subc As CSubscriber
      Dim Agr As CAgreement
      
      Set Acc = New CAccount
      Set Subc = New CSubscriber
      Set Agr = New CAgreement
      
      Acc.AddEditMode = ShowMode
      Subc.AddEditMode = ShowMode
      Agr.AddEditMode = ShowMode
      
      Acc.Flag = "A"
      Subc.Flag = "A"
      Agr.Flag = "A"
      
      Call Acc.ActSubs.add(Subc)
      Call Acc.ActAgrmnts.add(Agr)
      Call m_Customer.CstAccounts.add(Acc)
      
      Acc.ACCOUNT_NO = "DMY000"
      Acc.ACCOUNT_NAME = "DMY000"
      Acc.ACCOUNT_STATUS = -1
      Acc.ACCOUNT_TYPE = -1
      Acc.MASTER_FLAG = "Y"
      Acc.ENABLE_FLAG = "Y"
      
      Subc.DUMMY_FLAG = "Y"
      Subc.SUBSCRIBER_NO = "DMY999"
      Subc.SUBSCRIBER_STATUS = "Y"
      
      Agr.SOC_CODE = ""
      Agr.SOC_FEATURE_ID = -1
      Agr.SOC_ID = -1
      Agr.EXCLUDE_FLAG = "N"
      Agr.EFFECTIVE_DATE = -2
      Agr.EXPIRE_DATE = -1
      Agr.ISSUE_DATE = Now
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub
Private Function SetValue() As Long
   If ssoVolume.Value Then
      SetValue = 1
   ElseIf ssoRound.Value Then
      SetValue = 2
   Else
      SetValue = 1
   End If
End Function
Public Sub GetValue(ID As Long)
   If ID = 1 Then
      ssoVolume.Value = True
   ElseIf ID = 2 Then
      ssoRound.Value = True
   Else
      ssoVolume.Value = True
      ssoRound.Value = False
   End If
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   If ShowMode = SHOW_EDIT Then
         If Not VerifyAccessRight("MAIN_CUSTOMER_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   
   If Not VerifyTextControl(lblShortName, txtShortName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBusinessType, cboBusinessType, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblEnterpriseType, cboEnterpriseType, False) Then
      Exit Function
   End If


   If Not CheckUniqueNs(CUSTCODE_UNIQUE, txtShortName.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtShortName.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   m_Customer.AddEditMode = ShowMode
   m_Customer.BIRTH_DATE = -1
   m_Customer.CUSTOMER_PASSWORD = ""
   m_Customer.EMAIL = txtEmail.Text
   m_Customer.WEBSITE = txtWebSite.Text
   m_Customer.CUSTOMER_TYPE = cboBusinessType.ItemData(Minus2Zero(cboBusinessType.ListIndex))
   m_Customer.CUSTOMER_GRADE = cboEnterpriseType.ItemData(Minus2Zero(cboEnterpriseType.ListIndex))
   m_Customer.Credit = Val(txtCredit.Text)
   m_Customer.CUSTOMER_CODE = txtShortName.Text
   m_Customer.BUSINESS_DESC = txtBusinessDesc.Text
   m_Customer.NORMAL_DISCOUNT = Val(txtDiscountPercent.Text)
   m_Customer.RESPONSE_BY = uctlSaleByLookup.MyCombo.ItemData(Minus2Zero(uctlSaleByLookup.MyCombo.ListIndex))
   m_Customer.EXP_CODE = txtExpCode.Text
   m_Customer.CREDIT_LIMIT = Val(txtCreditLimit.Text)
    m_Customer.MAX_CREDIT = Val(txtMaxCredit.Text)
   m_Customer.SUSPEND_SALES = Check2Flag(chkSuspendSales.Value)
   m_Customer.FREE_PRICE_FLAG = Check2Flag(chkFreePriceFlag.Value)
   m_Customer.CHECK_CREDIT_FLAG = Check2Flag(chkCheckCreditFlag.Value)
   m_Customer.WEEK_CREDIT_LIMIT = Val(txtWeekCreditLimit.Text)
   m_Customer.CASH_FLAG = Check2Flag(chkCheckCashFlag.Value)
   m_Customer.CAL_PRICE_PART_CENTER_FLAG = Check2Flag(chkCalPricePartCenterFlag.Value)
   m_Customer.CAL_PRICE_DLC_CENTER_FLAG = Check2Flag(chkCalPriceDlcCenterFlag.Value)
   m_Customer.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   
   m_Customer.PRO_COMMISSION_BAG = Val(txtConBag1.Text)
   m_Customer.PRO_CHEER_BAG = Val(txtConBag2.Text)
   m_Customer.PRO_DST_BAG = Val(txtConBag3.Text)
   m_Customer.PRO_OTHER1_BAG = Val(txtConBag4.Text)
   m_Customer.PRO_OTHER2_BAG = Val(txtConBag5.Text)
   m_Customer.PRO_OTHER3_BAG = Val(txtConBag6.Text)
   
   m_Customer.PRO_COMMISSION_KG = Val(txtConKg1.Text)
   m_Customer.PRO_CHEER_KG = Val(txtConKg2.Text)
   m_Customer.PRO_DST_KG = Val(txtConKg3.Text)
   m_Customer.PRO_OTHER1_KG = Val(txtConKg4.Text)
   m_Customer.PRO_OTHER2_KG = Val(txtConKg5.Text)
   m_Customer.PRO_OTHER3_KG = Val(txtConKg6.Text)
   
   m_Customer.PRO_OTHER1_NAME = txtPRO_OTHER1_NAME.Text
   m_Customer.PRO_OTHER2_NAME = txtPRO_OTHER2_NAME.Text
   m_Customer.PRO_OTHER3_NAME = txtPRO_OTHER3_NAME.Text
   
   m_Customer.PRICE_THINK_TYPE = cboRateType.ItemData(Minus2Zero(cboRateType.ListIndex))
   m_Customer.CAL_RATE_DELIVERY_TYPE = SetValue


   'Create Dummy account
   If m_Customer.CstAccounts.Count <= 0 Then
      Dim Acc As CAccount
      
      Set Acc = New CAccount
      
      Acc.ACCOUNT_NO = m_Customer.CUSTOMER_CODE
      Acc.Flag = "A"
      
      Call m_Customer.CstAccounts.add(Acc)
      
      Set Acc = Nothing
   End If

   Dim CstName As CCustomerName
   If m_Customer.CstNames.Count <= 0 Then
      Set CstName = New CCustomerName
      CstName.Flag = "A"
      Call m_Customer.CstNames.add(CstName)
   Else
      Set CstName = m_Customer.CstNames.Item(1)
      CstName.Flag = "E"
   End If
   
   Dim NAME As CName
   If m_Customer.CstNames.Count <= 0 Then
      Set NAME = CstName.NAME
      NAME.LONG_NAME = txtName.Text
      NAME.SHORT_NAME = txtShortName.Text
      NAME.Flag = "A"
   Else
      Set NAME = CstName.NAME
      NAME.LONG_NAME = txtName.Text
      NAME.SHORT_NAME = txtShortName.Text
      NAME.Flag = "E"
   End If
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditCustomer(m_Customer, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub

Private Sub cboRateType_Change()
   m_HasModify = True
End Sub

Private Sub cboRateType_Click()
   m_HasModify = True
End Sub

Private Sub chkCalPriceDlcCenterFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCalPricePartCenterFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCheckCashFlag_Click(Value As Integer)
    m_HasModify = True
   If Not VerifyAccessRight("CREDIT_CASH-PAY", "สามารถอนุมัติการซื้อขายด้วยเงินสด") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
End Sub

Private Sub chkCheckCreditFlag_Click(Value As Integer)
m_HasModify = True
If Value = 1 Then
   chkSuspendSales.Enabled = True
   If Val(txtCreditLimit.Text) <= 0 Then
      txtCreditLimit.Text = 1
   End If
ElseIf Value = 0 Then
   chkSuspendSales.Enabled = False
   chkSuspendSales.Value = ssCBUnchecked
   If Val(txtCreditLimit.Text) = 1 Then
      txtCreditLimit.Text = 0
   End If
End If
End Sub

Private Sub chkFreePriceFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkFreePriceFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkSuspendSales_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkSuspendSales_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditCustomerAddress.TempCollection = m_Customer.CstAddr
      frmAddEditCustomerAddress.ShowMode = SHOW_ADD
      frmAddEditCustomerAddress.HeaderText = MapText("เพิ่มที่อยู่")
      Load frmAddEditCustomerAddress
      frmAddEditCustomerAddress.Show 1

      OKClick = frmAddEditCustomerAddress.OKClick

      Unload frmAddEditCustomerAddress
      Set frmAddEditCustomerAddress = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAddr)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Set frmAddEditCustomerAccount.TempCollection = m_Customer.CstAccounts
      frmAddEditCustomerAccount.ShowMode = SHOW_ADD
      frmAddEditCustomerAccount.HeaderText = MapText("เพิ่มบัญชีลูกค้า")
      Load frmAddEditCustomerAccount
      frmAddEditCustomerAccount.Show 1

      OKClick = frmAddEditCustomerAccount.OKClick

      Unload frmAddEditCustomerAccount
      Set frmAddEditCustomerAccount = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAccounts)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.AddMenu(glbGuiConfigs.PictureItems)
      Set oMenu = Nothing
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      
      If lMenuChosen = HEAD_ACCOUNT Then
         Set frmAddEditCustomerPicture.ParentForm = Me
         Set frmAddEditCustomerPicture.TempCollection = m_Customer.CstPicture
         frmAddEditCustomerPicture.ShowMode = SHOW_ADD
         frmAddEditCustomerPicture.PictureType = HEAD_ACCOUNT
         frmAddEditCustomerPicture.HeaderText = MapText("เพิ่ม ") & PictureTypeToText(HEAD_ACCOUNT)
         Load frmAddEditCustomerPicture
         frmAddEditCustomerPicture.Show 1
   
         OKClick = frmAddEditCustomerPicture.OKClick
   
         Unload frmAddEditCustomerPicture
         Set frmAddEditCustomerPicture = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstPicture)
         GridEX1.Rebind
      End If
      
      End If
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      Set frmAddEditCustomerAccList.ParentForm = Me
      Set frmAddEditCustomerAccList.TempCollection = m_Customer.CstAccountList1s
      frmAddEditCustomerAccList.ShowMode = SHOW_ADD
      frmAddEditCustomerAccList.AccountListType = 1
      frmAddEditCustomerAccList.HeaderText = MapText("เพิ่มบัญชีสินค้า")
      Load frmAddEditCustomerAccList
      frmAddEditCustomerAccList.Show 1

      OKClick = frmAddEditCustomerAccList.OKClick

      Unload frmAddEditCustomerAccList
      Set frmAddEditCustomerAccList = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAccountList1s)
         GridEX1.Rebind
      End If
            
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      Set frmAddEditCustomerAccList.ParentForm = Me
      Set frmAddEditCustomerAccList.TempCollection = m_Customer.CstAccountList2s
      frmAddEditCustomerAccList.ShowMode = SHOW_ADD
      frmAddEditCustomerAccList.AccountListType = 2
      frmAddEditCustomerAccList.HeaderText = MapText("เพิ่มบัญชีบริการ")
      Load frmAddEditCustomerAccList
      frmAddEditCustomerAccList.Show 1

      OKClick = frmAddEditCustomerAccList.OKClick

      Unload frmAddEditCustomerAccList
      Set frmAddEditCustomerAccList = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAccountList2s)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      Set frmAddEditCustomerAccList.ParentForm = Me
      Set frmAddEditCustomerAccList.TempCollection = m_Customer.CstAccountList3s
      frmAddEditCustomerAccList.ShowMode = SHOW_ADD
      frmAddEditCustomerAccList.AccountListType = 3
      frmAddEditCustomerAccList.HeaderText = MapText("เพิ่มบัญชีธนาคาร")
      Load frmAddEditCustomerAccList
      frmAddEditCustomerAccList.Show 1

      OKClick = frmAddEditCustomerAccList.OKClick

      Unload frmAddEditCustomerAccList
      Set frmAddEditCustomerAccList = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAccountList3s)
         GridEX1.Rebind
      End If
      
      ElseIf TabStrip1.SelectedItem.Index = 7 Then
      Set frmAddEditCustomerAccFol.ParentForm = Me
      Set frmAddEditCustomerAccFol.TempCollection = m_Customer.CstAccFol
      frmAddEditCustomerAccFol.ShowMode = SHOW_ADD
  '    frmAddEditCustomerAccFol.AccountListType = 3
      frmAddEditCustomerAccFol.HeaderText = MapText("เพิ่มรายละเอียดการติดตามฝ่ายบัญชี")
      Load frmAddEditCustomerAccFol
      frmAddEditCustomerAccFol.Show 1

      OKClick = frmAddEditCustomerAccFol.OKClick

      Unload frmAddEditCustomerAccFol
      Set frmAddEditCustomerAccFol = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAccFol)
         GridEX1.Rebind
      End If
      
      
      ElseIf TabStrip1.SelectedItem.Index = 8 Then
      Set frmAddEditCustomerMKTFol.ParentForm = Me
      Set frmAddEditCustomerMKTFol.TempCollection = m_Customer.CstMKTFol
      frmAddEditCustomerMKTFol.ShowMode = SHOW_ADD
  '    frmAddEditCustomerAccFol.AccountListType = 3
      frmAddEditCustomerMKTFol.HeaderText = MapText("เพิ่มรายละเอียดการติดตามฝ่ายขาย")
      Load frmAddEditCustomerMKTFol
      frmAddEditCustomerMKTFol.Show 1

      OKClick = frmAddEditCustomerMKTFol.OKClick

      Unload frmAddEditCustomerMKTFol
      Set frmAddEditCustomerMKTFol = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstMKTFol)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 9 Then
      Set frmAddEditCustomerFreelance.ParentForm = Me
      Set frmAddEditCustomerFreelance.TempCollection = m_Customer.CstFreelance
      frmAddEditCustomerFreelance.ShowMode = SHOW_ADD
      frmAddEditCustomerFreelance.HeaderText = MapText("เพิ่มรายละเอียดการติดตามฝ่ายขาย")
      Load frmAddEditCustomerFreelance
      frmAddEditCustomerFreelance.Show 1

      OKClick = frmAddEditCustomerFreelance.OKClick

      Unload frmAddEditCustomerFreelance
      Set frmAddEditCustomerFreelance = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstFreelance)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 10 Then
      Set frmAddEditDeliveryCus.ParentForm = Me
      Set frmAddEditDeliveryCus.TempCollection = m_Customer.CstdDeliveryCus
      frmAddEditDeliveryCus.ShowMode = SHOW_ADD
      frmAddEditDeliveryCus.HeaderText = MapText("เพิ่มข้อมูลสถานที่จัดส่ง")
      Load frmAddEditDeliveryCus
      frmAddEditDeliveryCus.Show 1

      OKClick = frmAddEditDeliveryCus.OKClick

      Unload frmAddEditDeliveryCus
      Set frmAddEditDeliveryCus = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstdDeliveryCus)
         GridEX1.Rebind
      End If
'   ElseIf TabStrip1.SelectedItem.Index = 11 Then
'      Set frmAddEditPromotional.ParentForm = Me
'      Set frmAddEditPromotional.TempCollection = m_Customer.CstdPromotion
'      frmAddEditPromotional.ShowMode = SHOW_ADD
'      frmAddEditPromotional.HeaderText = MapText("เพิ่มข้อมูลส่งเสริมการขาย")
'      Load frmAddEditPromotional
'      frmAddEditPromotional.Show 1
'
'      OKClick = frmAddEditPromotional.OKClick
'
'      Unload frmAddEditPromotional
'      Set frmAddEditPromotional = Nothing
'
'      If OKClick Then
'         GridEX1.ItemCount = CountItem(m_Customer.CstdPromotion)
'         GridEX1.Rebind
'      End If
   End If
   'frmAddEditPromotional
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdAuto_Click()
Dim No As String

   If Trim(txtShortName.Text) = "" Then
      Call glbDatabaseMngr.GenerateNumber(CUSTOMER_NUMBER, No, glbErrorLog)
      txtShortName.Text = No
   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_Customer.CstAddr.Remove (ID2)
      Else
         m_Customer.CstAddr.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Customer.CstAddr)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_Customer.CstAccounts.Item(ID2).MASTER_FLAG = "Y" Then
         glbErrorLog.LocalErrorMsg = "ไม่สามารถลบบัญชีพื้นฐานได้"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   
      If ID1 <= 0 Then
         m_Customer.CstAccounts.Remove (ID2)
      Else
         m_Customer.CstAccounts.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Customer.CstAccounts)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      If ID1 <= 0 Then
         m_Customer.CstPicture.Remove (ID2)
      Else
         m_Customer.CstPicture.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Customer.CstPicture)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If ID1 <= 0 Then
         m_Customer.CstAccountList1s.Remove (ID2)
      Else
         m_Customer.CstAccountList1s.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Customer.CstAccountList1s)
      GridEX1.Rebind
      m_HasModify = True
      
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      If ID1 <= 0 Then
         m_Customer.CstAccountList2s.Remove (ID2)
      Else
         m_Customer.CstAccountList2s.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Customer.CstAccountList2s)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      If ID1 <= 0 Then
         m_Customer.CstAccountList3s.Remove (ID2)
      Else
         m_Customer.CstAccountList3s.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Customer.CstAccountList3s)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 7 Then
      If ID1 <= 0 Then
         m_Customer.CstAccFol.Remove (ID2)
      Else
         m_Customer.CstAccFol.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Customer.CstAccFol)
      GridEX1.Rebind
      m_HasModify = True
      
      
   ElseIf TabStrip1.SelectedItem.Index = 8 Then
      If ID1 <= 0 Then
         m_Customer.CstMKTFol.Remove (ID2)
      Else
         m_Customer.CstMKTFol.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Customer.CstMKTFol)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 9 Then
      If ID1 <= 0 Then
         m_Customer.CstFreelance.Remove (ID2)
      Else
         m_Customer.CstFreelance.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Customer.CstFreelance)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 10 Then
      If ID1 <= 0 Then
         m_Customer.CstdDeliveryCus.Remove (ID2)
      Else
         m_Customer.CstdDeliveryCus.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Customer.CstdDeliveryCus)
      GridEX1.Rebind
      m_HasModify = True
'    ElseIf TabStrip1.SelectedItem.Index = 11 Then
'      If ID1 <= 0 Then
'         m_Customer.CstdPromotion.Remove (ID2)
'      Else
'         m_Customer.CstdPromotion.Item(ID2).Flag = "D"
'      End If
'
'      GridEX1.ItemCount = CountItem(m_Customer.CstdPromotion)
'      GridEX1.Rebind
'      m_HasModify = True
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim ID2 As Long
Dim OKClick As Boolean
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   ID2 = Val(GridEX1.Value(1))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditCustomerAddress.ID = ID
      Set frmAddEditCustomerAddress.TempCollection = m_Customer.CstAddr
      frmAddEditCustomerAddress.HeaderText = MapText("แก้ไขที่อยู่")
      frmAddEditCustomerAddress.ShowMode = SHOW_EDIT
      Load frmAddEditCustomerAddress
      frmAddEditCustomerAddress.Show 1

      OKClick = frmAddEditCustomerAddress.OKClick

      Unload frmAddEditCustomerAddress
      Set frmAddEditCustomerAddress = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAddr)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditCustomerAccount.ID = ID
      Set frmAddEditCustomerAccount.TempCollection = m_Customer.CstAccounts
      frmAddEditCustomerAccount.HeaderText = MapText("แก้ไขบัญชีลูกค้า")
      frmAddEditCustomerAccount.ShowMode = SHOW_EDIT
      Load frmAddEditCustomerAccount
      frmAddEditCustomerAccount.Show 1

      OKClick = frmAddEditCustomerAccount.OKClick

      Unload frmAddEditCustomerAccount
      Set frmAddEditCustomerAccount = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAccounts)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      Set frmAddEditCustomerPicture.ParentForm = Me
      frmAddEditCustomerPicture.ID = ID
      Set frmAddEditCustomerPicture.TempCollection = m_Customer.CstPicture
      frmAddEditCustomerPicture.ShowMode = SHOW_EDIT
      frmAddEditCustomerPicture.PictureType = HEAD_ACCOUNT
      frmAddEditCustomerPicture.HeaderText = MapText("แก้ไข ") & PictureTypeToText(HEAD_ACCOUNT)
      Load frmAddEditCustomerPicture
      frmAddEditCustomerPicture.Show 1

      OKClick = frmAddEditCustomerPicture.OKClick

      Unload frmAddEditCustomerPicture
      Set frmAddEditCustomerPicture = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstPicture)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      frmAddEditCustomerAccList.ID = ID
      Set frmAddEditCustomerAccList.ParentForm = Me
      Set frmAddEditCustomerAccList.TempCollection = m_Customer.CstAccountList1s
      frmAddEditCustomerAccList.ShowMode = SHOW_EDIT
      frmAddEditCustomerAccList.AccountListType = 1
      frmAddEditCustomerAccList.HeaderText = MapText("แก้ไขบัญชีสินค้า")
      Load frmAddEditCustomerAccList
      frmAddEditCustomerAccList.Show 1

      OKClick = frmAddEditCustomerAccList.OKClick

      Unload frmAddEditCustomerAccList
      Set frmAddEditCustomerAccList = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAccountList1s)
         GridEX1.Rebind
      End If
            
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      frmAddEditCustomerAccList.ID = ID
      Set frmAddEditCustomerAccList.ParentForm = Me
      Set frmAddEditCustomerAccList.TempCollection = m_Customer.CstAccountList2s
      frmAddEditCustomerAccList.ShowMode = SHOW_EDIT
      frmAddEditCustomerAccList.AccountListType = 2
      frmAddEditCustomerAccList.HeaderText = MapText("แก้ไขบัญชีบริการ")
      Load frmAddEditCustomerAccList
      frmAddEditCustomerAccList.Show 1

      OKClick = frmAddEditCustomerAccList.OKClick

      Unload frmAddEditCustomerAccList
      Set frmAddEditCustomerAccList = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAccountList2s)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      frmAddEditCustomerAccList.ID = ID
      Set frmAddEditCustomerAccList.ParentForm = Me
      Set frmAddEditCustomerAccList.TempCollection = m_Customer.CstAccountList3s
      frmAddEditCustomerAccList.ShowMode = SHOW_EDIT
      frmAddEditCustomerAccList.AccountListType = 3
      frmAddEditCustomerAccList.HeaderText = MapText("แก้ไขบัญชีธนาคาร")
      Load frmAddEditCustomerAccList
      frmAddEditCustomerAccList.Show 1

      OKClick = frmAddEditCustomerAccList.OKClick

      Unload frmAddEditCustomerAccList
      Set frmAddEditCustomerAccList = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAccountList3s)
         GridEX1.Rebind
      End If


    ElseIf TabStrip1.SelectedItem.Index = 7 Then
      frmAddEditCustomerAccFol.ID = ID
      Set frmAddEditCustomerAccFol.ParentForm = Me
      Set frmAddEditCustomerAccFol.TempCollection = m_Customer.CstAccFol
      frmAddEditCustomerAccFol.ShowMode = SHOW_EDIT
'      frmAddEditCustomerAccFol.AccountListType = 3
      frmAddEditCustomerAccFol.HeaderText = MapText("แก้ไขรายละเอียดการติดตามฝ่ายบัญชี")
      Load frmAddEditCustomerAccFol
      frmAddEditCustomerAccFol.Show 1

      OKClick = frmAddEditCustomerAccFol.OKClick

      Unload frmAddEditCustomerAccFol
      Set frmAddEditCustomerAccFol = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAccFol)
         GridEX1.Rebind
      End If
      
      ElseIf TabStrip1.SelectedItem.Index = 8 Then
      frmAddEditCustomerMKTFol.ID = ID
      Set frmAddEditCustomerMKTFol.ParentForm = Me
      Set frmAddEditCustomerMKTFol.TempCollection = m_Customer.CstMKTFol
      frmAddEditCustomerMKTFol.ShowMode = SHOW_EDIT
'      frmAddEditCustomerAccFol.AccountListType = 3
      frmAddEditCustomerMKTFol.HeaderText = MapText("แก้ไขรายละเอียดการติดตามฝ่ายขาย")
      Load frmAddEditCustomerMKTFol
      frmAddEditCustomerMKTFol.Show 1

      OKClick = frmAddEditCustomerMKTFol.OKClick

      Unload frmAddEditCustomerMKTFol
      Set frmAddEditCustomerMKTFol = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstMKTFol)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 9 Then
      frmAddEditCustomerFreelance.ID = ID
      Set frmAddEditCustomerFreelance.ParentForm = Me
      Set frmAddEditCustomerFreelance.TempCollection = m_Customer.CstFreelance
      frmAddEditCustomerFreelance.ShowMode = SHOW_EDIT
      frmAddEditCustomerFreelance.HeaderText = MapText("แก้ไขมรายละเอียดการติดตามฝ่ายขาย")
      Load frmAddEditCustomerFreelance
      frmAddEditCustomerFreelance.Show 1

      OKClick = frmAddEditCustomerFreelance.OKClick

      Unload frmAddEditCustomerFreelance
      Set frmAddEditCustomerFreelance = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstFreelance)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 10 Then
      frmAddEditDeliveryCus.ID = ID
      frmAddEditDeliveryCus.ID2 = ID2
      Set frmAddEditDeliveryCus.ParentForm = Me
      Set frmAddEditDeliveryCus.TempCollection = m_Customer.CstdDeliveryCus
      frmAddEditDeliveryCus.ShowMode = SHOW_EDIT
      frmAddEditDeliveryCus.HeaderText = MapText("แก้ไขสถานที่จัดส่ง")
      Load frmAddEditDeliveryCus
      frmAddEditDeliveryCus.Show 1

      OKClick = frmAddEditDeliveryCus.OKClick

      Unload frmAddEditDeliveryCus
      Set frmAddEditDeliveryCus = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstdDeliveryCus)
         GridEX1.Rebind
      End If
'   ElseIf TabStrip1.SelectedItem.Index = 11 Then
'      frmAddEditDeliveryCus.ID = ID
'      Set frmAddEditPromotional.ParentForm = Me
'      Set frmAddEditPromotional.TempCollection = m_Customer.CstdPromotion
'      frmAddEditPromotional.ShowMode = SHOW_EDIT
'      frmAddEditPromotional.HeaderText = MapText("แก้ไขข้อมูลการส่งเสริมการขาย")
'      Load frmAddEditPromotional
'      frmAddEditPromotional.Show 1
'
'      OKClick = frmAddEditPromotional.OKClick
'
'      Unload frmAddEditPromotional
'      Set frmAddEditPromotional = Nothing
'
'      If OKClick Then
'         GridEX1.ItemCount = CountItem(m_Customer.CstdPromotion)
'         GridEX1.Rebind
'      End If
   End If
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdEditCon_Click()
'   frmVerifyAccRight.AccName = "CREDIT_PROMOTIONAL"
'   frmVerifyAccRight.AccDesc = "สามารถเปลี่ยนแปลงเงื่อนไขส่งเสริมการขาย"
'   Load frmVerifyAccRight
'   frmVerifyAccRight.Show 1
'
'   If frmVerifyAccRight.GrantRight Then
'      Unload frmVerifyAccRight
'      Set frmVerifyAccRight = Nothing
'   Else
'      Unload frmVerifyAccRight
'      Set frmVerifyAccRight = Nothing
'      Exit Sub
'   End If
'
'   SSFrame2.Enabled = True
'   SSFrame3.Enabled = True
'   lblRateType.Enabled = True
'   cboRateType.Enabled = True
End Sub

Private Sub cmdEditCredit_Click()
   frmVerifyAccRight.AccName = "CREDIT_CONTROL"
   frmVerifyAccRight.AccDesc = "สามารถเปลี่ยนแปลงเครดิตซื้อขาย"
   Load frmVerifyAccRight
   frmVerifyAccRight.Show 1
   
   If frmVerifyAccRight.GrantRight Then
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
   Else
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
      Exit Sub
   End If
            
   txtCreditLimit.Enabled = True
   chkSuspendSales.Enabled = True
   chkFreePriceFlag.Enabled = True
   chkCheckCreditFlag.Enabled = True
   chkCheckCashFlag.Enabled = True
   chkCalPricePartCenterFlag.Enabled = True
   chkCalPriceDlcCenterFlag.Enabled = True
   txtMaxCredit.Enabled = True
   txtWeekCreditLimit.Enabled = True
   
   SSFrame2.Enabled = True
   SSFrame3.Enabled = True
   cboRateType.Enabled = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadCustomerType(cboBusinessType)
      Call LoadCustomerGrade(cboEnterpriseType)

      
      Call LoadEmployee(uctlSaleByLookup.MyCombo, m_Employees)
      Set uctlSaleByLookup.MyCollection = m_Employees
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2)
      Set uctlLocationLookup.MyCollection = m_Locations
      
      Call InitDoRateType2(cboRateType)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Customer.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_Customer.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
'   SSFrame2.Width = ScaleWidth - 2 * GridEX1.Left
'   SSFrame2.HEIGHT = ScaleHeight - GridEX1.Top - 620
   TabStrip1.Width = GridEX1.Width
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdEditCredit.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Customer = Nothing
   Set m_Employees = Nothing
   Set m_Customers = Nothing
   Set m_Locations = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 11550
   Col.Caption = MapText("ที่อยู่")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 0
   Col.Caption = MapText("ที่อยู่ไม่แสดงเบอร์โทร")
End Sub

Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 1470
   Col.Caption = MapText("เลขที่บัญชี")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 6855
   Col.Caption = MapText("รายละเอียด")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 3240
   Col.Caption = MapText("แพคเกจ")
End Sub
Private Sub InitGrid3()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2355
   Col.Caption = MapText("ประเภทเอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 8855
   Col.Caption = MapText("ที่อยู่ที่จัดเก็บ")

End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblWebsite, MapText("เว็บไซต์"))
   Call InitNormalLabel(lblShortName, MapText("รหัสลูกค้า"))
   Call InitNormalLabel(lblEnterpriseType, MapText("ระดับลูกค้า"))
   Call InitNormalLabel(lblName, MapText("ชื่อลูกค้า"))
   Call InitNormalLabel(lblEmail, MapText("อีเมลล์"))
   Call InitNormalLabel(lblBusinessType, MapText("ประเภทลูกค้า"))
   Call InitNormalLabel(lblBusinessDesc, MapText("รายละเอียดลูกค้า"))
   Call InitNormalLabel(lblCredit, MapText("เครดิต"))
   Call InitNormalLabel(lblDiscountPercent, MapText("% ส่วนลด"))
   Call InitNormalLabel(lblResponseBy, MapText("ผู้รับผิดชอบ"))
   Call InitNormalLabel(lblLocationLookup, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(lblExpCode, MapText("รหัส EXP."))
   Call InitNormalLabel(lblCreditLimit, MapText("วงเงิน"))
   Call InitNormalLabel(lblMaxCredit, MapText("เครดิตสูงสุด"))
   Call InitNormalLabel(Label1, MapText("วัน"))
   Call InitNormalLabel(lblWeekCreditLimit, MapText("วงเงินสัปดาห์"))
   
   Call InitCheckBox(chkSuspendSales, "ระงับการขาย")
   Call InitCheckBox(chkCheckCreditFlag, "เช็คเครดิต")
   Call InitCheckBox(chkCheckCashFlag, "จ่ายเงินสด")
   Call InitCheckBox(chkCalPricePartCenterFlag, "คิดค่าสินค้าจากส่วนกลาง")
   Call InitCheckBox(chkCalPriceDlcCenterFlag, "คิดค่าขนส่งจากส่วนกลาง")
   Call InitCheckBox(chkFreePriceFlag, "ไม่คิดราคาขาย")


   Call InitOptionEx(ssoVolume, "คิดตามปริมาณ")
   Call InitOptionEx(ssoRound, "คิดตามเที่ยว")

   Call InitNormalFrame(SSFrame2, "เงื่อนไขส่งเสริมการขาย/ปริมาณ")
   Call InitNormalFrame(SSFrame3, "เงื่อนไขการคิดค่าขนส่ง")
   Call InitNormalLabel(lblBag, MapText("/ถุง(30 กก.)"))
   Call InitNormalLabel(lblKg, MapText("/ก.ก."))
   Call InitNormalLabel(lblCon1, MapText("ค่าคอม"))
   Call InitNormalLabel(lblCon2, MapText("ค่าเชียร์"))
   Call InitNormalLabel(lblCon3, MapText("ค่าทอย"))
   Call InitNormalLabel(lblCon4, MapText("อื่นๆ1"))
   Call InitNormalLabel(lblCon5, MapText("อื่นๆ2"))
   Call InitNormalLabel(lblCon6, MapText("อื่นๆ3"))
   Call InitNormalLabel(lblBath1, MapText("บาท"))
   Call InitNormalLabel(lblBath2, MapText("บาท"))
   Call InitNormalLabel(lblBath3, MapText("บาท"))
   Call InitNormalLabel(lblBath4, MapText("บาท"))
   Call InitNormalLabel(lblBath5, MapText("บาท"))
   Call InitNormalLabel(lblBath6, MapText("บาท"))
   
   Call InitNormalLabel(lblRateType, MapText("คิดราคาแบบ"))
   
   Call txtConBag1.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtConBag2.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtConBag3.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtConBag4.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtConBag5.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtConBag6.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtConKg1.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtConKg2.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtConKg3.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtConKg4.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtConKg5.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtConKg6.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtPRO_OTHER1_NAME.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtPRO_OTHER2_NAME.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtPRO_OTHER3_NAME.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)

   Call InitCombo(cboBusinessType)
   Call InitCombo(cboEnterpriseType)
   Call InitCombo(cboRateType)
   
   Call txtShortName.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtEmail.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtWebSite.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBusinessDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtCredit.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtExpCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtCreditLimit.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtWeekCreditLimit.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEditCredit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEditCon.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
'   SSFrame2.BackColor = LoadPicture(glbParameterObj.NormalForm1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdEditCredit, MapText("แก้ไข CREDIT"))
   Call InitMainButton(cmdEditCon, MapText("แก้ไข"))
   'cmdEditCon
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("ที่อยู่")
   TabStrip1.Tabs.add().Caption = MapText("บัญชีลูกค้า")
   TabStrip1.Tabs.add().Caption = MapText("เอกสารรูปภาพ")
   TabStrip1.Tabs.add().Caption = MapText("บัญชีสินค้า")
   TabStrip1.Tabs.add().Caption = MapText("บัญชีบริการ")
   TabStrip1.Tabs.add().Caption = MapText("บัญชีธนาคาร")
   TabStrip1.Tabs.add().Caption = MapText("การติดตามฝ่ายบัญชี")
   TabStrip1.Tabs.add().Caption = MapText("การติดตามฝ่ายขาย")
   TabStrip1.Tabs.add().Caption = MapText("ข้อมูลฟรีแลนซ์")
   TabStrip1.Tabs.add().Caption = MapText("ข้อมูลสถานที่จัดส่ง")
'   TabStrip1.Tabs.add().Caption = MapText("ข้อมูลส่งเสริมขาย")
   
   frmVerifyAccRight.AccName = "CREDIT_CONTROL"
   frmVerifyAccRight.AccDesc = "สามารถเปลี่ยนแปลงเครดิตซื้อขาย"
   
   If Not VerifyAccessRight("CREDIT_CONTROL", "", 2) Then
      txtCreditLimit.Enabled = False
      chkSuspendSales.Enabled = False
      chkCheckCreditFlag.Enabled = False
      txtMaxCredit.Enabled = False
      txtWeekCreditLimit.Enabled = False
      chkCheckCashFlag.Enabled = False
      chkCalPricePartCenterFlag.Enabled = False
      chkCalPriceDlcCenterFlag.Enabled = False
      chkFreePriceFlag.Enabled = False
      
      cboRateType.Enabled = False
      SSFrame2.Enabled = False
      SSFrame3.Enabled = False
   Else
      txtCreditLimit.Enabled = True
      chkSuspendSales.Enabled = False
      chkFreePriceFlag.Enabled = False
      chkCheckCreditFlag.Enabled = True
      txtMaxCredit.Enabled = True
      txtWeekCreditLimit.Enabled = True
      chkCheckCashFlag.Enabled = True
      chkCalPricePartCenterFlag.Enabled = True
      chkCalPriceDlcCenterFlag.Enabled = True
      
      cboRateType.Enabled = True
      SSFrame2.Enabled = True
      SSFrame3.Enabled = True
    End If
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Customer = New CCustomer
   Set m_Employees = New Collection
   Set m_Customers = New Collection
   Set m_Locations = New Collection
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim IsOK As Boolean
Dim OKClick As Boolean
Dim ReportFlag As Boolean
Dim ReportKey As String
Dim Report As CReportInterface
Dim Rc As CReportConfig
Dim iCount As Long
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long
   
   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   If TabStrip1.SelectedItem.Index <> 1 Then
      Exit Sub
   End If
   If m_HasModify Or (m_Customer.CUSTOMER_ID <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   TempID1 = GridEX1.Value(1)
   
   ReportMode = 1
   ReportFlag = False
   
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("พิมพ์ที่อยู่+เบอร์โทร", "พิมพ์ที่อยู่", "ตั้งค่า")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      ReportKey = "CReportMain004"
      
      Set Report = New CReportMain004
      ReportFlag = True
      Call Report.AddParam(GridEX1.Value(3), "ADDRESS")
   ElseIf lMenuChosen = 2 Then
      ReportKey = "CReportMain004"
      
      Set Report = New CReportMain004
      ReportFlag = True
      Call Report.AddParam(GridEX1.Value(4), "ADDRESS")
   ElseIf lMenuChosen = 3 Then
      ReportKey = "CReportMain004"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("จดหมาย ชื่อที่อยู่ลูกค้า")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   End If
      
   If Not Report Is Nothing Then
      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
      Call Report.AddParam(m_Customer.CUSTOMER_ID, "CUSTOMER_ID")
      Call Report.AddParam(m_Customer.CUSTOMER_NAME, "CUSTOMER_NAME")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(MapText("จดหมาย"), "REPORT_HEADER")
   End If
   
   Call EnableForm(Me, False)
   If ReportFlag Then
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = ""
      Load frmReport
      frmReport.Show 1
         
      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ReportMode = ReportMode
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
      Load frmReportConfig
      frmReportConfig.Show 1
      
      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   Call EnableForm(Me, True)

End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_Customer.CstAddr Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CCustomerAddress
      Dim Addr As CAddress
      If m_Customer.CstAddr.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Customer.CstAddr, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      Set Addr = CR.Addresses

      Values(1) = Addr.ADDRESS_ID
      Values(2) = RealIndex
      Values(3) = Addr.PackAddressEx1(True)
      Values(4) = Addr.PackAddressEx1(False)
   
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_Customer.CstAccounts Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ca As CAccount
      If m_Customer.CstAccounts.Count <= 0 Then
         Exit Sub
      End If
      Set Ca = GetItem(m_Customer.CstAccounts, RowIndex, RealIndex)
      If Ca Is Nothing Then
         Exit Sub
      End If

      Values(1) = Ca.ACCOUNT_ID
      Values(2) = RealIndex
      Values(3) = Ca.ACCOUNT_NO
      Values(4) = Ca.NOTE
      Values(5) = Ca.ActAgrmnts(1).SOC_CODE
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      If m_Customer.CstPicture Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Cs As CCustomerPicture
      If m_Customer.CstPicture.Count <= 0 Then
         Exit Sub
      End If
      Set Cs = GetItem(m_Customer.CstPicture, RowIndex, RealIndex)
      If Cs Is Nothing Then
         Exit Sub
      End If

      Values(1) = Cs.GetFieldValue("CUSTOMER_PICTURE_ID")
      Values(2) = RealIndex
      Values(3) = PictureTypeToText(Cs.GetFieldValue("CUSTOMER_PICTURE_TYPE"))
      Values(4) = Cs.GetFieldValue("CUSTOMER_PICTURE_PATH")
   
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   
      If m_Customer.CstAccountList1s Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Acl1 As CCustomerAccountList
      If m_Customer.CstAccountList1s.Count <= 0 Then
         Exit Sub
      End If
      Set Acl1 = GetItem(m_Customer.CstAccountList1s, RowIndex, RealIndex)
      If Acl1 Is Nothing Then
         Exit Sub
      End If

      Values(1) = Acl1.GetFieldValue("CUSTOMER_ACCOUNT_LIST_ID")
      Values(2) = RealIndex
      Values(3) = Acl1.GetFieldValue("PART_GROUP_NAME")
      Values(4) = Acl1.GetFieldValue("DEBIT_NO")
      Values(5) = Acl1.GetFieldValue("CREDIT_NO")
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   
      If m_Customer.CstAccountList2s Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Acl2 As CCustomerAccountList
      If m_Customer.CstAccountList2s.Count <= 0 Then
         Exit Sub
      End If
      Set Acl2 = GetItem(m_Customer.CstAccountList2s, RowIndex, RealIndex)
      If Acl2 Is Nothing Then
         Exit Sub
      End If

      Values(1) = Acl2.GetFieldValue("CUSTOMER_ACCOUNT_LIST_ID")
      Values(2) = RealIndex
      Values(3) = Acl2.GetFieldValue("FEATURE_TYPE_NAME")
      Values(4) = Acl2.GetFieldValue("DEBIT_NO")
      Values(5) = Acl2.GetFieldValue("CREDIT_NO")
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
   
      If m_Customer.CstAccountList3s Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Acl3 As CCustomerAccountList
      If m_Customer.CstAccountList3s.Count <= 0 Then
         Exit Sub
      End If
      Set Acl3 = GetItem(m_Customer.CstAccountList3s, RowIndex, RealIndex)
      If Acl3 Is Nothing Then
         Exit Sub
      End If

      Values(1) = Acl3.GetFieldValue("CUSTOMER_ACCOUNT_LIST_ID")
      Values(2) = RealIndex
      Values(3) = Acl3.GetFieldValue("BANK_ACCOUNT_NAME")
      Values(4) = Acl3.GetFieldValue("DEBIT_NO")
      Values(5) = Acl3.GetFieldValue("CREDIT_NO")
      
      
      ElseIf TabStrip1.SelectedItem.Index = 7 Then
      If m_Customer.CstAccFol Is Nothing Then
         Exit Sub
      End If
      
      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim AccFol As CAccFol
      If m_Customer.CstAccFol.Count <= 0 Then
         Exit Sub
      End If
      Set AccFol = GetItem(m_Customer.CstAccFol, RowIndex, RealIndex)
      If AccFol Is Nothing Then
         Exit Sub
      End If

      Values(1) = AccFol.ACC_FOL_ID
      Values(2) = RealIndex
      Values(3) = DateToStringExtEx2(AccFol.FOL_DATE)
      Values(4) = AccFol.CANCEL_FLAG
      Values(5) = AccFol.FOL_NOTE
      
     ElseIf TabStrip1.SelectedItem.Index = 8 Then
      If m_Customer.CstMKTFol Is Nothing Then
         Exit Sub
      End If
      
      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim MKTFol As CMKTFol
      If m_Customer.CstMKTFol.Count <= 0 Then
         Exit Sub
      End If
      Set MKTFol = GetItem(m_Customer.CstMKTFol, RowIndex, RealIndex)
      If MKTFol Is Nothing Then
         Exit Sub
      End If

      Values(1) = MKTFol.MKT_FOL_ID
      Values(2) = RealIndex
      Values(3) = DateToStringExtEx2(MKTFol.FOL_DATE)
      Values(4) = MKTFol.CANCEL_FLAG
      Values(5) = MKTFol.FOL_NOTE
   ElseIf TabStrip1.SelectedItem.Index = 9 Then
      If m_Customer.CstFreelance Is Nothing Then
         Exit Sub
      End If
      
      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim FreelanceItem As CFreelanceItem
      If m_Customer.CstFreelance.Count <= 0 Then
         Exit Sub
      End If
      Set FreelanceItem = GetItem(m_Customer.CstFreelance, RowIndex, RealIndex)
      If FreelanceItem Is Nothing Then
         Exit Sub
      End If

      Values(1) = FreelanceItem.FREELANCE_ID
      Values(2) = RealIndex
      Values(3) = FreelanceItem.FREELANCE_CODE
      Values(4) = FreelanceItem.FREELANCE_NAME & " " & FreelanceItem.FREELANCE_LASTNAME
    ElseIf TabStrip1.SelectedItem.Index = 10 Then
      If m_Customer.CstdDeliveryCus Is Nothing Then
         Exit Sub
      End If
      
      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim DC As CDeliveryCus
      If m_Customer.CstdDeliveryCus.Count <= 0 Then
         Exit Sub
      End If
      Set DC = GetItem(m_Customer.CstdDeliveryCus, RowIndex, RealIndex)
      If DC Is Nothing Then
         Exit Sub
      End If

      Values(1) = DC.DELIVERY_CUS_ITEM_ID
      Values(2) = RealIndex
      Values(3) = DC.DELIVERY_CUS_ITEM_CODE
      Values(4) = DC.DELIVERY_CUS_ITEM_NAME
'   ElseIf TabStrip1.SelectedItem.Index = 11 Then
'      If m_Customer.CstdPromotion Is Nothing Then
'         Exit Sub
'      End If
'
'      If RowIndex <= 0 Then
'         Exit Sub
'      End If
'
'      Dim Pt As CPromotional
'      If m_Customer.CstdPromotion.Count <= 0 Then
'         Exit Sub
'      End If
'      Set Pt = GetItem(m_Customer.CstdPromotion, RowIndex, RealIndex)
'      If Pt Is Nothing Then
'         Exit Sub
'      End If
'
'      Values(1) = Pt.PROMOTIONAL_ITEM_ID
'      Values(2) = RealIndex
'      Values(3) = Pt.PROMOTIONAL_DETAIL_ID
'      Values(4) = Pt.PROMOTIONAL_DETAIL_NAME
'      Values(5) = Pt.PROMOTIONAL_RATE
'      Values(6) = Pt.UNIT_TYPE
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub sscRound_Click(Value As Integer)
    m_HasModify = True
End Sub

Private Sub sscVolume_Click(Value As Integer)
    m_HasModify = True
End Sub

Private Sub ssoRound_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub ssoVolume_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub TabStrip1_Click()
   GridEX1.Top = 5600 '5010
   GridEX1.Left = 150
   GridEX1.Visible = False

   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_Customer.CstAddr)
      GridEX1.Rebind
      GridEX1.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid2
      GridEX1.ItemCount = CountItem(m_Customer.CstAccounts)
      GridEX1.Rebind
      GridEX1.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      Call InitGrid3
      GridEX1.ItemCount = CountItem(m_Customer.CstPicture)
      GridEX1.Rebind
      GridEX1.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      Call InitGrid4
      GridEX1.ItemCount = CountItem(m_Customer.CstAccountList1s)
      GridEX1.Rebind
      GridEX1.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      Call InitGrid5
      GridEX1.ItemCount = CountItem(m_Customer.CstAccountList2s)
      GridEX1.Rebind
      GridEX1.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      Call InitGrid6
      GridEX1.ItemCount = CountItem(m_Customer.CstAccountList3s)
      GridEX1.Rebind
      GridEX1.Visible = True
    ElseIf TabStrip1.SelectedItem.Index = 7 Then
      Call InitGrid7
      GridEX1.ItemCount = CountItem(m_Customer.CstAccFol)
      GridEX1.Rebind
      GridEX1.Visible = True
     ElseIf TabStrip1.SelectedItem.Index = 8 Then
      Call InitGrid8
      GridEX1.ItemCount = CountItem(m_Customer.CstMKTFol)
      GridEX1.Rebind
      GridEX1.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 9 Then
      Call InitGrid9
      GridEX1.ItemCount = CountItem(m_Customer.CstFreelance)
      GridEX1.Rebind
      GridEX1.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 10 Then
      Call InitGrid10
      GridEX1.ItemCount = CountItem(m_Customer.CstdDeliveryCus)
      GridEX1.Rebind
      GridEX1.Visible = True
'   ElseIf TabStrip1.SelectedItem.Index = 11 Then
'      Call InitGrid11
'      GridEX1.ItemCount = CountItem(m_Customer.CstdPromotion)
'      GridEX1.Rebind
'      GridEX1.Visible = True
   End If
End Sub

Private Sub txtBusinessDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtConBag1_Change()
   m_HasModify = True
End Sub

Private Sub txtConBag2_Change()
   m_HasModify = True
End Sub

Private Sub txtConBag3_Change()
   m_HasModify = True
End Sub

Private Sub txtConBag4_Change()
   m_HasModify = True
End Sub

Private Sub txtConBag5_Change()
   m_HasModify = True
End Sub

Private Sub txtConBag6_Change()
   m_HasModify = True
End Sub

Private Sub txtConKg1_Change()
   m_HasModify = True
End Sub

Private Sub txtConKg2_Change()
   m_HasModify = True
End Sub

Private Sub txtConKg3_Change()
   m_HasModify = True
End Sub

Private Sub txtConKg4_Change()
   m_HasModify = True
End Sub

Private Sub txtConKg5_Change()
   m_HasModify = True
End Sub

Private Sub txtConKg6_Change()
   m_HasModify = True
End Sub

Private Sub txtCredit_Change()
   m_HasModify = True
   If Val(txtCredit.Text) = 0 Then
      txtDiscountPercent.Enabled = True
   Else
      txtDiscountPercent.Enabled = False
   End If
End Sub
Private Sub txtCreditLimit_Change()
   m_HasModify = True
End Sub
Private Sub txtDiscountPercent_Change()
   m_HasModify = True
End Sub

Private Sub txtEmail_Change()
   m_HasModify = True
End Sub

Private Sub txtExpCode_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxCredit_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtPRO_OTHER1_NAME_Change()
   m_HasModify = True
End Sub

Private Sub txtPRO_OTHER2_NAME_Change()
   m_HasModify = True
End Sub

Private Sub txtPRO_OTHER3_NAME_Change()
   m_HasModify = True
End Sub

Private Sub txtShortName_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtWebSite_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextLookup1_Change()
   m_HasModify = True
End Sub

Private Sub txtWeekCreditLimit_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlSaleByLookup_Change()
   m_HasModify = True
End Sub
Public Sub RefreshGrid()
   GridEX1.ItemCount = CountItem(m_Customer.CstPicture)
   GridEX1.Rebind
End Sub
Public Sub RefreshGridAccountList(AccountListType As Long)
   If AccountListType = 1 Then
      GridEX1.ItemCount = CountItem(m_Customer.CstAccountList1s)
   ElseIf AccountListType = 2 Then
      GridEX1.ItemCount = CountItem(m_Customer.CstAccountList2s)
   ElseIf AccountListType = 3 Then
      GridEX1.ItemCount = CountItem(m_Customer.CstAccountList3s)
   End If
   GridEX1.Rebind
End Sub

Private Sub InitGrid4()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2355
   Col.Caption = MapText("กลุ่มวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 4400
   Col.Caption = MapText("DEBIT")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 4455
   Col.Caption = MapText("CREDIT")
   
End Sub
Private Sub InitGrid5()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2355
   Col.Caption = MapText("ประเภทบริการ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 4400
   Col.Caption = MapText("DEBIT")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 4455
   Col.Caption = MapText("CREDIT")
   
End Sub
Private Sub InitGrid6()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2355
   Col.Caption = MapText("บัญชีหมายเลข")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 4400
   Col.Caption = MapText("DEBIT")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 4455
   Col.Caption = MapText("CREDIT")
 End Sub
   
Private Sub InitGrid7()
   Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2355
   Col.Caption = MapText("วันที่")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2000
   Col.Caption = MapText("สถานะลูกค้า")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 6000
   Col.Caption = MapText("รายละเอียดประวัติลูกค้า")
   

End Sub
Private Sub InitGrid8()
   Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2355
   Col.Caption = MapText("วันที่")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2000
   Col.Caption = MapText("สถานะลูกค้า")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 6000
   Col.Caption = MapText("รายละเอียดประวัติลูกค้า")
   

End Sub
Private Sub InitGrid9()
   Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2355
   Col.Caption = MapText("รหัสฟรีแลนซ์")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 5000
   Col.Caption = MapText("ชื่อฟรีแลนซ์")

'   Set Col = GridEX1.Columns.add '4
'   Col.Width = 6000
'   Col.Caption = MapText("รายละเอียดประวัติลูกค้า")
   

End Sub
Private Sub InitGrid10()
   Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2355
   Col.Caption = MapText("รหัสสถานที่")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 5000
   Col.Caption = MapText("ชื่อสถานที่จัดส่ง")
End Sub

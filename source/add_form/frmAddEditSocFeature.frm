VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditSocFeature 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditSocFeature.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8505
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15002
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlFeatureLookup 
         Height          =   405
         Left            =   2520
         TabIndex        =   2
         Top             =   2040
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2745
         Left            =   90
         TabIndex        =   37
         Top             =   5040
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   4842
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin GridEX20.GridEX GridEX1 
            Height          =   2715
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   4789
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            MethodHoldFields=   -1  'True
            AllowColumnDrag =   0   'False
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            HeaderFontName  =   "AngsanaUPC"
            FontSize        =   12
            ColumnHeaderHeight=   480
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   2
            Column(1)       =   "frmAddEditSocFeature.frx":27A2
            Column(2)       =   "frmAddEditSocFeature.frx":286A
            FormatStylesCount=   5
            FormatStyle(1)  =   "frmAddEditSocFeature.frx":290E
            FormatStyle(2)  =   "frmAddEditSocFeature.frx":2A6A
            FormatStyle(3)  =   "frmAddEditSocFeature.frx":2B1A
            FormatStyle(4)  =   "frmAddEditSocFeature.frx":2BCE
            FormatStyle(5)  =   "frmAddEditSocFeature.frx":2CA6
            ImageCount      =   0
            PrinterProperties=   "frmAddEditSocFeature.frx":2D5E
         End
      End
      Begin VB.ComboBox cboRateType 
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
         Left            =   8280
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3360
         Width           =   2685
      End
      Begin VB.ComboBox cboUnit 
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
         Left            =   8280
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   2685
      End
      Begin prjFarmManagement.uctlTextBox txtSocCode 
         Height          =   435
         Left            =   2520
         TabIndex        =   0
         Top             =   1140
         Width           =   4005
         _ExtentX        =   11615
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   30
         TabIndex        =   22
         Top             =   7800
         Width           =   12240
         _ExtentX        =   21590
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   8490
            TabIndex        =   18
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditSocFeature.frx":2F36
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   10140
            TabIndex        =   19
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdEdit 
            Height          =   525
            Left            =   1680
            TabIndex        =   16
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   525
            Left            =   60
            TabIndex        =   15
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditSocFeature.frx":3250
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   525
            Left            =   3330
            TabIndex        =   17
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditSocFeature.frx":356A
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   21
         Top             =   0
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtOC 
         Height          =   435
         Left            =   2520
         TabIndex        =   3
         Top             =   2460
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRC 
         Height          =   435
         Left            =   8280
         TabIndex        =   4
         Top             =   2460
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAC 
         Height          =   435
         Left            =   2520
         TabIndex        =   8
         Top             =   3360
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMinimum 
         Height          =   435
         Left            =   2520
         TabIndex        =   12
         Top             =   4260
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRate 
         Height          =   435
         Left            =   8280
         TabIndex        =   11
         Top             =   3810
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRoundingFactor 
         Height          =   435
         Left            =   8280
         TabIndex        =   13
         Top             =   4260
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPackageRate 
         Height          =   435
         Left            =   2520
         TabIndex        =   10
         Top             =   3810
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlFeatureTypeLookup 
         Height          =   405
         Left            =   2520
         TabIndex        =   1
         Top             =   1590
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtLogisticWage 
         Height          =   435
         Left            =   2520
         TabIndex        =   5
         Top             =   2910
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox uctlTextBox2 
         Height          =   435
         Left            =   8280
         TabIndex        =   6
         Top             =   2910
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin VB.Label lblLogisticWage 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   690
         TabIndex        =   44
         Top             =   3030
         Width           =   1665
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6240
         TabIndex        =   43
         Top             =   3060
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   4350
         TabIndex        =   42
         Top             =   3030
         Width           =   1065
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   10110
         TabIndex        =   41
         Top             =   3030
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblFeatureType 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   40
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblPackageRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   690
         TabIndex        =   39
         Top             =   3930
         Width           =   1695
      End
      Begin VB.Label lblBath5 
         Height          =   315
         Left            =   4350
         TabIndex        =   38
         Top             =   3930
         Width           =   705
      End
      Begin VB.Label lblRoundingFactor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6180
         TabIndex        =   36
         Top             =   4380
         Width           =   1995
      End
      Begin VB.Label lblRateType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6510
         TabIndex        =   35
         Top             =   3450
         Width           =   1695
      End
      Begin VB.Label lblBath4 
         Height          =   315
         Left            =   10050
         TabIndex        =   34
         Top             =   3900
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblUC 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6450
         TabIndex        =   33
         Top             =   3930
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblMinimum 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1110
         TabIndex        =   32
         Top             =   4380
         Width           =   1305
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6630
         TabIndex        =   31
         Top             =   300
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblBath3 
         Height          =   315
         Left            =   4350
         TabIndex        =   30
         Top             =   3480
         Width           =   1065
      End
      Begin VB.Label lblBath2 
         Height          =   315
         Left            =   10110
         TabIndex        =   29
         Top             =   2580
         Width           =   1065
      End
      Begin VB.Label lblBath1 
         Height          =   315
         Left            =   4350
         TabIndex        =   28
         Top             =   2580
         Width           =   1065
      End
      Begin VB.Label lblAC 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   480
         TabIndex        =   27
         Top             =   3480
         Width           =   1905
      End
      Begin VB.Label lblRC 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6240
         TabIndex        =   26
         Top             =   2610
         Width           =   1905
      End
      Begin VB.Label lblOC 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   690
         TabIndex        =   25
         Top             =   2580
         Width           =   1665
      End
      Begin VB.Label lblFeatureCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   630
         TabIndex        =   24
         Top             =   2130
         Width           =   1755
      End
      Begin VB.Label lblSocCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   510
         TabIndex        =   23
         Top             =   1230
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmAddEditSocFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Features As Collection
Private m_FeatureTypes As Collection
Private m_TempCol As Collection
Private m_FeatureCode As String
Private m_FeatureDesc As String
Private m_TempFeatures As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public TempCollection As Collection
Public SocCode As String
Public SocID As Long
Public SocPartType As Long

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.ColumnHeaderFont.Size = 16
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = MapText("ID1")

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = MapText("ID2")

   Set Col = GridEX1.Columns.add '3
   Col.Width = 3030
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จาก")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 3300
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ถึง")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2475
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ความกว้าง")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2475
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคา/หน่วย")
End Sub

Private Sub CopyStepTierVolume(D As Collection)
Dim C As CStpTierVol
Dim Tmp As CStpTierVol

   For Each C In D
      Set Tmp = New CStpTierVol
      Call C.CopyObject(Tmp)
      Tmp.Flag = "I"
      Call m_TempCol.add(Tmp)
      Set Tmp = Nothing
   Next C
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim D As CSocFeature

   If Flag Then
      Call EnableForm(Me, False)
      
      Set D = TempCollection.Item(id)
      Call CopyStepTierVolume(D.StepTierVol)
      
      txtSocCode.Text = SocCode
      If D.OC_FLAG = "Y" Then
         txtOC.Text = D.OCRate.RATE_AMOUNT
      End If
      If D.RC_FLAG = "Y" Then
         txtRC.Text = D.RCRate.RATE_AMOUNT
      End If
      If D.AC_FLAG = "Y" Then
         txtAC.Text = D.ACRate.RATE_AMOUNT
      End If
      If D.UC_FLAG = "Y" Then
         txtRate.Text = D.UCRate.RATE_AMOUNT
         txtPackageRate.Text = D.UCRate.PKG_RATE_AMOUNT
         txtLogisticWage.Text = D.UCRate.LOG_RATE_AMOUNT
      End If
      If D.MINIMUM_FLAG = "Y" Then
         txtMinimum.Text = D.MINIMUM_UNIT
      End If
      If SocPartType = 1 Then
         uctlFeatureTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlFeatureTypeLookup.MyCombo, D.FEATURE_TYPE)
         uctlFeatureLookup.MyCombo.ListIndex = IDToListIndex(uctlFeatureLookup.MyCombo, D.FEATURE_ID)
      ElseIf SocPartType = 3 Then
         uctlFeatureTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlFeatureTypeLookup.MyCombo, D.PART_TYPE)
         uctlFeatureLookup.MyCombo.ListIndex = IDToListIndex(uctlFeatureLookup.MyCombo, D.PART_ITEM_ID)
      End If
      cboRateType.ListIndex = IDToListIndex(cboRateType, D.RATE_TYPE)
      txtRoundingFactor.Text = D.ROUNDING_FACTOR
      m_FeatureCode = D.FEATURE_CODE
      m_FeatureDesc = D.FEATURE_DESC
   
      GridEX1.ItemCount = CountItem(m_TempCol)
      GridEX1.Rebind
      
      Call EnableForm(Me, True)
   End If
   
   If ItemCount > 0 Then
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboFeatureLevel_Click()
   m_HasModify = True
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Sf As CSocFeature

   If Not VerifyCombo(lblFeatureCode, uctlFeatureLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblOC, txtOC, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblRC, txtRC, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblAC, txtAC, True) Then
      Exit Function
   End If
   If Not VerifyCombo(lblRateType, cboRateType, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMinimum, txtMinimum, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblUC, txtRate, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblRoundingFactor, txtRoundingFactor, True) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If ShowMode = SHOW_ADD Then
      Set Sf = New CSocFeature

      Sf.Flag = "A"
      Sf.RCRate.Flag = "A"
      Sf.OCRate.Flag = "A"
      Sf.ACRate.Flag = "A"
      Sf.UCRate.Flag = "A"
      
      Call TempCollection.add(Sf)
   Else
      Set Sf = TempCollection(id)
      
      Sf.Flag = "E"
      Sf.RCRate.Flag = "E"
      Sf.OCRate.Flag = "E"
      Sf.ACRate.Flag = "E"
      Sf.UCRate.Flag = "E"
   End If
   
   Sf.RTTYPE_NAME = cboRateType.Text
   Sf.RATE_TYPE = cboRateType.ItemData(Minus2Zero(cboRateType.ListIndex))
   If SocPartType = 1 Then
      Sf.FEATURE_ID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
      Sf.FEATURE_TYPE = uctlFeatureTypeLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureTypeLookup.MyCombo.ListIndex))
      Sf.FEATURE_CODE = uctlFeatureLookup.MyTextBox.Text
      Sf.FEATURE_DESC = uctlFeatureLookup.MyCombo.Text
   ElseIf SocPartType = 3 Then
      Sf.PART_ITEM_ID = uctlFeatureLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureLookup.MyCombo.ListIndex))
      Sf.PART_TYPE = uctlFeatureTypeLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureTypeLookup.MyCombo.ListIndex))
      Sf.PART_NO = uctlFeatureLookup.MyTextBox.Text
      Sf.PART_DESC = uctlFeatureLookup.MyCombo.Text
   End If
   Sf.ROUNDING_FACTOR = Val(txtRoundingFactor.Text)
   Sf.USE_END_FLAG = "N"
   Sf.USE_START_FLAG = "N"

   If Trim(txtOC.Text) = "" Then
      Sf.OC_FLAG = "N"
   Else
      Sf.OC_FLAG = "Y"
      Sf.OCRate.RATE_AMOUNT = Val(txtOC.Text)
   End If
   If Trim(txtRC.Text) = "" Then
      Sf.RC_FLAG = "N"
      Sf.RCRate.RATE_AMOUNT = 0
   Else
      Sf.RC_FLAG = "Y"
      Sf.RCRate.RATE_AMOUNT = Val(txtRC.Text)
   End If
   If Trim(txtAC.Text) = "" Then
      Sf.AC_FLAG = "N"
      Sf.ACRate.RATE_AMOUNT = 0
   Else
      Sf.AC_FLAG = "Y"
      Sf.ACRate.RATE_AMOUNT = Val(txtAC.Text)
   End If
   If (Trim(txtRate.Text) = "") And (Trim(txtPackageRate.Text) = "") Then
      Sf.UC_FLAG = "N"
      Sf.UCRate.RATE_AMOUNT = 0
      Sf.UCRate.PKG_RATE_AMOUNT = 0
      Sf.UCRate.LOG_RATE_AMOUNT = 0
   Else
      Sf.UC_FLAG = "Y"
      Sf.UCRate.RATE_AMOUNT = Val(txtRate.Text)
      Sf.UCRate.PKG_RATE_AMOUNT = Val(txtPackageRate.Text)
      Sf.UCRate.LOG_RATE_AMOUNT = Val(txtLogisticWage.Text)
   End If
   
   If Trim(txtMinimum.Text) = "" Then
      Sf.MINIMUM_FLAG = "N"
      Sf.MINIMUM_UNIT = 0
   Else
      Sf.MINIMUM_FLAG = "Y"
      Sf.MINIMUM_UNIT = Val(txtMinimum.Text)
   End If
   
   Set Sf.StepTierVol = Nothing
   Set Sf.StepTierVol = New Collection
   Set Sf.StepTierVol = m_TempCol
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cboRateType_Click()
Dim RateTypeID As Long
Dim TempID As Long
Dim Sp As CSystemParam
Dim D As CStpTierVol

   RateTypeID = RATE_FLAT
    TempID = cboRateType.ItemData(Minus2Zero(cboRateType.ListIndex))
    If TempID = RateTypeID Then
        GridEX1.Enabled = False
        cmdAdd.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        txtRate.Enabled = True
    
         For Each D In m_TempCol
            D.Flag = "D"
         Next D
         GridEX1.ItemCount = CountItem(m_TempCol)
         GridEX1.Rebind
    Else
        GridEX1.Enabled = True
        cmdAdd.Enabled = True
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
        txtRate.Enabled = False
    End If
    
    m_HasModify = True
End Sub

Private Sub cboRateType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboUnit_Click()
   m_HasModify = True
End Sub

Private Sub chkStartEndFlag_Click()
   m_HasModify = True
End Sub

Private Sub ReArrangeStepTier()
Dim C As CStpTierVol
Dim NextFrom As Double

   NextFrom = 0
   For Each C In m_TempCol
      If C.Flag <> "D" Then
         C.FROM_QUANTITY = NextFrom
         C.TO_QUANTITY = C.FROM_QUANTITY + C.Width
         NextFrom = C.TO_QUANTITY

         If C.Flag <> "A" Then
            C.Flag = "E"
         End If
      End If
   Next C
End Sub

Private Function GetMaxFrom() As CStpTierVol
Dim C As CStpTierVol
Dim Tmp As CStpTierVol
Dim MAX As Double

   MAX = -1
   For Each C In m_TempCol
      If C.Flag <> "D" Then
         If C.FROM_QUANTITY > MAX Then
            MAX = C.FROM_QUANTITY
            Set Tmp = C
         End If
      End If
   Next C
   
   Set GetMaxFrom = Tmp
End Function

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim Tmp As CStpTierVol
Dim FromQ As Double

   Set Tmp = GetMaxFrom
   If Tmp Is Nothing Then
      FromQ = 0
   Else
      FromQ = Tmp.TO_QUANTITY
   End If
    frmAddEditStepTier.SocPartType = SocPartType
    frmAddEditStepTier.ShowMode = SHOW_ADD
    frmAddEditStepTier.FROM_QUANTITY = FromQ
    frmAddEditStepTier.HeaderText = "เพิ่มช่วงราคา"
    Set frmAddEditStepTier.TempCollection = m_TempCol
    Load frmAddEditStepTier
    frmAddEditStepTier.Show 1
    
   OKClick = frmAddEditStepTier.OKClick
   
   Unload frmAddEditStepTier
    Set frmAddEditStepTier = Nothing

   If OKClick Then
  Dim T As CStpTierVol
  For Each T In m_TempCol
   ''Debug.Print T.FROM_QUANTITY & " " & T.TO_QUANTITY & " " & T.Width & " " & T.RATE_AMOUNT
  Next T
      Call ReArrangeStepTier
      GridEX1.ItemCount = CountItem(m_TempCol)
      GridEX1.Rebind
      
      m_HasModify = True
   Else
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
   
'   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_TempCol.Remove (ID2)
      Else
         m_TempCol.Item(ID2).Flag = "D"
      End If

      Call ReArrangeStepTier
      GridEX1.ItemCount = CountItem(m_TempCol)
      GridEX1.Rebind
      m_HasModify = True
'   End If
End Sub

Private Sub cmdEdit_Click()
Dim OKClick As Boolean
Dim Tmp As CStpTierVol
Dim FromQ As Double
Dim id As Long
Dim ID2 As Long
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   id = Val(GridEX1.Value(2))
   ID2 = Val(GridEX1.Value(1))
   OKClick = False
   
   frmAddEditStepTier.id = id
    frmAddEditStepTier.ShowMode = SHOW_EDIT
    frmAddEditStepTier.HeaderText = MapText("แก้ไขช่วงราคา")
    Set frmAddEditStepTier.TempCollection = m_TempCol
    Load frmAddEditStepTier
    frmAddEditStepTier.Show 1
    
   OKClick = frmAddEditStepTier.OKClick
   
   Unload frmAddEditStepTier
    Set frmAddEditStepTier = Nothing

   If OKClick Then
      Call ReArrangeStepTier
      GridEX1.ItemCount = CountItem(m_TempCol)
      GridEX1.Rebind
      m_HasModify = True
   Else
   End If
End Sub

Private Sub cmdOK_Click()

   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
Dim Sp As CSystemParam
Dim FeatureTypeID As Long

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      If SocPartType = 1 Then
         Call LoadFeatureType(uctlFeatureTypeLookup.MyCombo, m_FeatureTypes)
         Set uctlFeatureTypeLookup.MyCollection = m_FeatureTypes
'         Call LoadFeature(uctlFeatureLookup.MyCombo, m_TempFeatures)
'         Set uctlFeatureLookup.MyCollection = m_TempFeatures
      ElseIf SocPartType = 3 Then
         Call LoadPartType(uctlFeatureTypeLookup.MyCombo, m_FeatureTypes)
         Set uctlFeatureTypeLookup.MyCollection = m_FeatureTypes
'         Call LoadPartItem(uctlFeatureLookup.MyCombo, m_TempFeatures)
'         Set uctlFeatureLookup.MyCollection = m_TempFeatures
      End If
      
      Call InitRateType(cboRateType)
      txtSocCode.Text = SocCode
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call QueryData(True)
      Else
         cboRateType.ListIndex = IDToListIndex(cboRateType, RATE_FLAT)
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
   ElseIf Shift = 1 And KeyCode = 112 Then
      If glbUser.EXCEPTION_FLAG = "Y" Then
         glbUser.EXCEPTION_FLAG = "N"
      Else
         glbUser.EXCEPTION_FLAG = "Y"
      End If
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_Features = Nothing
   Set m_TempCol = Nothing
   Set m_TempFeatures = Nothing
   Set m_FeatureTypes = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitFormLayout()
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   Me.KeyPreview = True
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Call InitHeaderFooter(pnlHeader, pnlFooter)
      
   Call txtSocCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call InitNormalLabel(lblSocCode, MapText("แพคเกจ"))
   txtSocCode.Enabled = False
            
'   Call InitCheckBox(chkStartEndFlag, MapText("ต้องการเวลาสิ้นสุด"))
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   Call txtAC.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtRC.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtOC.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtMinimum.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtRate.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtPackageRate.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtRoundingFactor.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtLogisticWage.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtOC.Enabled = False
   txtRC.Enabled = False
   If SocPartType = 1 Then
      Call InitNormalLabel(lblFeatureType, MapText("ประเภทสินค้า/บริการ"))
      Call InitNormalLabel(lblFeatureCode, MapText("รหัสสินค้า/บริการ"))
      Call InitNormalLabel(lblUC, MapText("ค่าบริการ"))
      Call InitNormalLabel(lblPackageRate, MapText("ค่าบริการ/ถุง"))
   Else
      Call InitNormalLabel(lblFeatureType, MapText("ประเภทสินค้า/วัตถุดิบ"))
      Call InitNormalLabel(lblFeatureCode, MapText("รหัสสินค้า/วัตถุดิบ"))
      Call InitNormalLabel(lblUC, MapText("ค่าสินค้า/Kg"))
      Call InitNormalLabel(lblPackageRate, MapText("ค่าสินค้า/ถุง"))
   End If
   Call InitNormalLabel(lblLogisticWage, MapText("ค่าจ้างขนส่ง/ถุง"))
   Call InitNormalLabel(lblOC, MapText("ค่าบริการแรกเข้า"))
   Call InitNormalLabel(lblRC, MapText("ค่าบริการรายรอบ"))
   Call InitNormalLabel(lblAC, MapText("ค่าบริการพิเศษ"))
   Call InitNormalLabel(lblBath1, MapText("บาท"))
   Call InitNormalLabel(lblBath2, MapText("บาท"))
   Call InitNormalLabel(lblBath3, MapText("บาท"))
   Call InitNormalLabel(lblBath4, MapText("บาท"))
   Call InitNormalLabel(lblBath5, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(lblUnit, MapText("หน่วยนับ"))
   Call InitNormalLabel(lblMinimum, MapText("จำนวนขั้นต่ำ"))
   Call InitNormalLabel(lblRateType, MapText("การคิดราคา"))
   Call InitNormalLabel(lblRoundingFactor, MapText("ตัวประกอบปัดเศษ"))
   
   Call InitCombo(cboUnit)
   Call SetEnableDisableComboBox(cboUnit, False)
   Call InitCombo(cboRateType)
   
   Call InitGrid1
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
   Set m_Features = New Collection
   Set m_TempCol = New Collection
   Set m_TempFeatures = New Collection
   Set m_FeatureTypes = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub


Private Sub txtFeatureCode_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

'   If TabStrip1.SelectedItem.Index = 1 Then
      If m_TempCol Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CStpTierVol
      If m_TempCol.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_TempCol, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.STPTIER_VOL_ID
      Values(2) = RealIndex
      Values(3) = FormatNumber(CR.FROM_QUANTITY)
      Values(4) = FormatNumber(CR.TO_QUANTITY)
      Values(5) = FormatNumber(CR.TO_QUANTITY - CR.FROM_QUANTITY)
      Values(6) = FormatNumber(CR.RATE_AMOUNT)
'   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub txtAC_Change()
   m_HasModify = True
End Sub

Private Sub txtLogisticWage_Change()
   m_HasModify = True
End Sub

Private Sub txtMinimum_Change()
   m_HasModify = True
End Sub

Private Sub txtOC_Change()
   m_HasModify = True
End Sub

Private Sub txtPackageRate_Change()
   m_HasModify = True
End Sub

Private Sub txtRate_Change()
   m_HasModify = True
End Sub

Private Sub txtRC_Change()
   m_HasModify = True
End Sub

Private Sub txtRoundingFactor_Change()
   m_HasModify = True
End Sub

Private Sub txtSocCode_Change()
   m_HasModify = True
End Sub

Private Sub uctlExpireDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTime1_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTime2_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlFeatureLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlFeatureTypeLookup_Change()
Dim PartTypeID As Long

   m_HasModify = True

   If SocPartType = 1 Then
      PartTypeID = uctlFeatureTypeLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureTypeLookup.MyCombo.ListIndex))
      If PartTypeID > 0 Then
         Call LoadFeature(uctlFeatureLookup.MyCombo, m_TempFeatures, PartTypeID)
         Set uctlFeatureLookup.MyCollection = m_TempFeatures
      End If
   ElseIf SocPartType = 3 Then
      PartTypeID = uctlFeatureTypeLookup.MyCombo.ItemData(Minus2Zero(uctlFeatureTypeLookup.MyCombo.ListIndex))
      If PartTypeID > 0 Then
         Call LoadPartItem(uctlFeatureLookup.MyCombo, m_TempFeatures, PartTypeID)
         Set uctlFeatureLookup.MyCollection = m_TempFeatures
      End If
   End If
End Sub

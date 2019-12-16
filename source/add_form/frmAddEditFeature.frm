VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditFeature 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmAddEditFeature.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2340
         Width           =   2955
      End
      Begin VB.ComboBox cboFeatureType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1920
         Width           =   2955
      End
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1470
         Width           =   4485
         _extentx        =   13309
         _extenty        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   10
         TabIndex        =   9
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   2955
         _extentx        =   5212
         _extenty        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3975
         Left            =   150
         TabIndex        =   14
         Top             =   3690
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   7011
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
         Column(1)       =   "frmAddEditFeature.frx":27A2
         Column(2)       =   "frmAddEditFeature.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditFeature.frx":290E
         FormatStyle(2)  =   "frmAddEditFeature.frx":2A6A
         FormatStyle(3)  =   "frmAddEditFeature.frx":2B1A
         FormatStyle(4)  =   "frmAddEditFeature.frx":2BCE
         FormatStyle(5)  =   "frmAddEditFeature.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditFeature.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtUnitWeight 
         Height          =   435
         Left            =   7800
         TabIndex        =   2
         Top             =   1470
         Visible         =   0   'False
         Width           =   1305
         _extentx        =   5212
         _extenty        =   767
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   17
         Top             =   3150
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
      Begin Threed.SSCheck chkBillDirectFlag 
         Height          =   345
         Left            =   4920
         TabIndex        =   19
         Top             =   2760
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "chkBillDirectFlag"
         TripleState     =   -1  'True
      End
      Begin Threed.SSCheck chkCancelFlag 
         Height          =   345
         Left            =   4920
         TabIndex        =   18
         Top             =   2340
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "chkCancelFlag"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblKg 
         Caption         =   "Label1"
         Height          =   345
         Left            =   9180
         TabIndex        =   16
         Top             =   1560
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblUnitWeight 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6360
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   2400
         Width           =   1575
      End
      Begin Threed.SSCheck chkPigFlag 
         Height          =   345
         Left            =   4950
         TabIndex        =   7
         Top             =   1950
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "chkPigFlag"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   60
         TabIndex        =   12
         Top             =   1980
         Width           =   1725
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   10
         Top             =   1110
         Width           =   1755
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5985
         TabIndex        =   6
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4335
         TabIndex        =   5
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFeature.frx":2F36
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Feature As CFeature
Private m_Sp As CSystemParam

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private Sub cmdPasswd_Click()

End Sub

Private Sub cboFeatureType_Click()
   m_HasModify = True
End Sub

Private Sub cboUnit_Click()
   m_HasModify = True
End Sub

Private Sub chkBillDirectFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCancelFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkPigFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me
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
   Col.Width = 3465
   Col.Caption = MapText("หมายเลขแพคเกจ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2745
   Col.Caption = MapText("ระดับแพคเกจ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 5355
   Col.Caption = MapText("ชื่อแพคเกจ")
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
   Col.Width = 6030
   Col.Caption = MapText("ชื่อซัพพลายเออร์")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2745
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดนำเข้ารวม")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2790
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("มูลค่านำเข้ารวม")
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_Feature.FEATURE_ID = ID
      m_Feature.QueryFlag = 1
      If Not glbDaily.QueryFeature(m_Feature, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Feature.PopulateFromRS(1, m_Rs)
      
      txtName.Text = m_Feature.FEATURE_DESC
      txtPartNo.Text = m_Feature.FEATURE_CODE
      cboFeatureType.ListIndex = IDToListIndex(cboFeatureType, m_Feature.FEATURE_TYPE)
      cboUnit.ListIndex = IDToListIndex(cboUnit, m_Feature.FEATURE_UNIT)
      chkPigFlag.Value = FlagToCheck(m_Feature.SERVICE_FLAG)
      chkCancelFlag.Value = FlagToCheck(m_Feature.FEATURE_STATUS)
      chkBillDirectFlag.Value = FlagToCheck(m_Feature.BILL_DIRECT_FLAG)
      TabStrip1_Click
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_Feature.SocFeatures Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CSocFeature
      If m_Feature.SocFeatures.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Feature.SocFeatures, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.SOC_FEATURE_ID
      Values(2) = RealIndex
      Values(3) = CR.SOC_CODE
      Values(4) = CR.SOC_SVCLVL_NAME
      Values(5) = CR.SOC_DESC
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("PACKAGE_FEATURE_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   If Not VerifyTextControl(lblPartNo, txtPartNo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblUnitWeight, txtUnitWeight, True) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPartType, cboFeatureType, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblUnit, cboUnit, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(FEATURENO_UNIQUE, txtName.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   m_Feature.FEATURE_ID = ID
   m_Feature.AddEditMode = ShowMode
   m_Feature.FEATURE_CODE = txtPartNo.Text
   m_Feature.FEATURE_DESC = txtName.Text
   m_Feature.FEATURE_TYPE = cboFeatureType.ItemData(Minus2Zero(cboFeatureType.ListIndex))
   m_Feature.FEATURE_UNIT = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))
   m_Feature.FEATURE_STATUS = "Y"
   m_Feature.SERVICE_FLAG = Check2Flag(chkPigFlag.Value)
   m_Feature.FEATURE_STATUS = Check2Flag(chkCancelFlag.Value)
   m_Feature.BILL_DIRECT_FLAG = Check2Flag(chkBillDirectFlag.Value)

   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditFeature(m_Feature, IsOK, True, glbErrorLog) Then
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadFeatureType(cboFeatureType)
      Call LoadUnit(cboUnit)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
      End If
      
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
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub InitFormLayout()
   Set m_Sp = GetSystemParam(glbSystemParams, "PROGRAM_OWNER")
   
   Call InitGrid1
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblName, MapText("ชื่อสินค้า/บริการ"))
   Call InitNormalLabel(lblPartNo, MapText("รหัสสินค้า/บริการ"))
   Call InitNormalLabel(lblPartType, MapText("ประเภทสินค้า"))
   Call InitNormalLabel(lblUnit, MapText("หน่วยวัด"))
   If m_Sp.PARAM_VALUE = "GLDSTK" Then
      Call InitNormalLabel(lblUnitWeight, MapText("น้ำหนัก/หน่วย"))
      Call InitNormalLabel(lblKg, MapText("กรัม"))
   Else
      Call InitNormalLabel(lblUnitWeight, MapText("ความหนาแน่น"))
      Call InitNormalLabel(lblKg, MapText(""))
   End If
   
   Call InitCheckBox(chkPigFlag, "งานบริการ")
   Call InitCheckBox(chkCancelFlag, "ใช้งาน")
   Call InitCheckBox(chkBillDirectFlag, "ออกบิลโดยตรง")

   
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtPartNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtUnitWeight.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboFeatureType)
   Call InitCombo(cboUnit)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("แพคเกจ")
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_Feature = New CFeature
   Set m_Rs = New ADODB.Recordset

   Call EnableForm(Me, False)
   m_HasActivate = False
      
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub txtLastName_Change()
   m_HasModify = True
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      
      GridEX1.ItemCount = CountItem(m_Feature.SocFeatures)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtPartNo_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub txtUnitWeight_Change()
   m_HasModify = True
End Sub

VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMigrate 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmMigrate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6641
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   3
         Top             =   1950
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   9
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   5
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   465
         Left            =   1860
         TabIndex        =   0
         Top             =   990
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtTableName 
         Height          =   465
         Left            =   1860
         TabIndex        =   2
         Top             =   1470
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   820
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSCommand cmdFile 
         Height          =   405
         Left            =   8310
         TabIndex        =   1
         Top             =   1020
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmMigrate.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblTableName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   14
         Top             =   1590
         Width           =   1575
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1860
         TabIndex        =   4
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmMigrate.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   13
         Top             =   2400
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   2430
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1110
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8505
         TabIndex        =   7
         Top             =   2880
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6855
         TabIndex        =   6
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmMigrate.frx":2DD6
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmMigrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Employee As CEmployee

Public id As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private Sub cmdPasswd_Click()

End Sub


Private Sub cboPartType_Click()
   m_HasModify = True
End Sub

Private Sub cboPosition_Click()
   m_HasModify = True
End Sub

Private Sub chkPigFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdFile_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.GDB)|*..GDB;*.GDB;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName.Text = dlgAdd.FileName
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

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_Employee.EMP_ID = id
      m_Employee.QueryFlag = 1
      If Not glbDaily.QueryEmployee(m_Employee, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Employee.PopulateFromRS(1, m_Rs)
      
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Employee.EMP_ID = id
   m_Employee.AddEditMode = ShowMode
   m_Employee.PASS_STATUS = "Y"
   
   m_Employee.EmpName.AddEditMode = ShowMode
   m_Employee.EName.AddEditMode = ShowMode
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditEmployee(m_Employee, IsOK, True, glbErrorLog) Then
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

Private Function MigrateUnit() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CUnit
Dim LFt As CLegacyUnit
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "UNIT"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyUnit
   LFt.UNIT_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CUnit
      Ft.AddEditMode = SHOW_ADD
      Ft.UNIT_ID = LFt.UNIT_ID
      Ft.UNIT_NO = LFt.UNIT_NAME
      Ft.UNIT_NAME = LFt.UNIT_NAME
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateUnit = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateUnit = False
End Function

Private Function MigrateCountry() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CCountry
Dim LFt As CLegacyCountry
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "COUNTRY"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyCountry
   LFt.COUNTRY_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CCountry
      Ft.AddEditMode = SHOW_ADD
      Ft.COUNTRY_ID = LFt.COUNTRY_ID
      Ft.CONTINENT_ID = LFt.CONTINENT_ID
      Ft.COUNTRY_NO = LFt.COUNTRY_NAME
      Ft.COUNTRY_NAME = LFt.COUNTRY_NAME
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateCountry = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateCountry = False
End Function

Private Function MigrateFeatureType() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CFeatureType
Dim LFt As CLegacyFeatureType
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "FEATURE_TYPE"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyFeatureType
   LFt.FEATYPE_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CFeatureType
      Ft.AddEditMode = SHOW_ADD
      Ft.FEATURE_TYPE_ID = LFt.FEATYPE_ID
      Ft.FEATURE_TYPE_NO = LFt.FEATYPE_CODE
      Ft.FEATURE_TYPE_NAME = LFt.FEATYPE_NAME
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateFeatureType = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateFeatureType = False
End Function

Private Function MigrateFeature() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CFeature
Dim LFt As CLegacyFeature
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "FEATURE"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyFeature
   LFt.FEATURE_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CFeature
      Ft.AddEditMode = SHOW_ADD
      Ft.FEATURE_ID = LFt.FEATURE_ID
      Ft.FEATURE_CODE = LFt.FEATURE_CODE
      Ft.FEATURE_DESC = LFt.FEATURE_DESC
      Ft.FEATURE_LEVEL = 0
      Ft.FEATURE_STATUS = "Y"
      Ft.FEATURE_TYPE = LFt.FEATURE_TYPE
      Ft.FEATURE_UNIT = LFt.FEATURE_UNIT
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateFeature = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateFeature = False
End Function

Private Function MigrateSoc() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CSoc
Dim LFt As CLegacySoc
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "SOC"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacySoc
   LFt.SOC_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CSoc
      Ft.AddEditMode = SHOW_ADD
      Ft.SOC_ID = LFt.SOC_ID
      Ft.SOC_CODE = LFt.SOC_CODE
      Ft.SOC_DESC = LFt.SOC_DESC
      Ft.SOC_LEVEL = LFt.SOC_LEVEL
      Ft.SOC_STATUS = 0
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateSoc = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateSoc = False
End Function

Private Function MigrateSocFeature() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CSocFeature
Dim LFt As CLegacySocFeature
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "SOC_FEATURE"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacySocFeature
   LFt.SOC_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CSocFeature
      Ft.AddEditMode = SHOW_ADD
      Ft.SOC_FEATURE_ID = LFt.SOC_FEATURE_ID
      Ft.SOC_ID = LFt.SOC_ID
      Ft.FEATURE_ID = LFt.FEATURE_ID
      Ft.PART_ITEM_ID = -1
      Ft.RC_FLAG = LFt.RC_FLAG
      Ft.UC_FLAG = LFt.UC_FLAG
      Ft.OC_FLAG = LFt.OC_FLAG
      Ft.AC_FLAG = LFt.AC_FLAG
      Ft.RATE_TYPE = LFt.RATE_TYPE
      Ft.MINIMUM_FLAG = LFt.MINIMUM_FLAG
      Ft.MINIMUM_UNIT = LFt.MINIMUM_UNIT
      Ft.USE_START_FLAG = LFt.USE_START_FLAG
      Ft.USE_END_FLAG = LFt.USE_END_FLAG
      Ft.ROUNDING_FACTOR = LFt.ROUNDING_FACTOR
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateSocFeature = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateSocFeature = False
End Function

Private Function MigrateStpTierVol() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CStpTierVol
Dim LFt As CLegacyStpTierVol
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "STPTIER_VOL"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyStpTierVol
   LFt.STPTIER_VOL_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CStpTierVol
      Ft.AddEditMode = SHOW_ADD
      Ft.STPTIER_VOL_ID = LFt.STPTIER_VOL_ID
      Ft.FROM_QUANTITY = LFt.FROM_QUANTITY
      Ft.TO_QUANTITY = LFt.TO_QUANTITY
      Ft.SOC_FEATURE_ID = LFt.SOC_FEATURE_ID
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateStpTierVol = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateStpTierVol = False
End Function

Private Function MigrateAcRate() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CAcRate
Dim LFt As CLegacyACRate
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "AC_RATE"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyACRate
   LFt.AC_RATE_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CAcRate
      Ft.AddEditMode = SHOW_ADD
      Ft.AC_RATE_ID = LFt.AC_RATE_ID
      Ft.RATE_AMOUNT = LFt.RATE_AMOUNT
      Ft.SOC_FEATURE_ID = LFt.SOC_FEATURE_ID
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateAcRate = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateAcRate = False
End Function

Private Function MigrateOcRate() As Boolean
On Error GoTo ErrorHandler
Dim Ft As COcRate
Dim LFt As CLegacyOCRate
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "OC_RATE"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyOCRate
   LFt.OC_RATE_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New COcRate
      Ft.AddEditMode = SHOW_ADD
      Ft.OC_RATE_ID = LFt.OC_RATE_ID
      Ft.RATE_AMOUNT = LFt.RATE_AMOUNT
      Ft.SOC_FEATURE_ID = LFt.SOC_FEATURE_ID
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateOcRate = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateOcRate = False
End Function

Private Function MigrateRcRate() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CRcRate
Dim LFt As CLegacyRCRate
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "RC_RATE"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyRCRate
   LFt.RC_RATE_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CRcRate
      Ft.AddEditMode = SHOW_ADD
      Ft.RC_RATE_ID = LFt.RC_RATE_ID
      Ft.RATE_AMOUNT = LFt.RATE_AMOUNT
      Ft.SOC_FEATURE_ID = LFt.SOC_FEATURE_ID
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateRcRate = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateRcRate = False
End Function

Private Function MigrateUcRate() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CUcRate
Dim LFt As CLegacyUCRate
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "UC_RATE"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyUCRate
   LFt.UC_RATE_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CUcRate
      Ft.AddEditMode = SHOW_ADD
      Ft.UC_RATE_ID = LFt.UC_RATE_ID
      Ft.RATE_AMOUNT = LFt.RATE_AMOUNT
      Ft.SOC_FEATURE_ID = LFt.SOC_FEATURE_ID
      Ft.STPTIER_VOL_ID = LFt.STPTIER_VOL_ID
      Ft.NULL_FLAG = LFt.NULL_FLAG
      Ft.STPTIER_VOL_ID = LFt.STPTIER_VOL_ID
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateUcRate = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateUcRate = False
End Function

Private Function MigrateName() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CName
Dim LFt As CLegacyName
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "NAME"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyName
   LFt.NAME_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CName
      Ft.AddEditMode = SHOW_ADD
      Ft.NAME_ID = LFt.NAME_ID
      Ft.NICK_NAME = LFt.NICK_NAME
      Ft.LANGUAGE_ID = LFt.LANGUAGE_ID
      Ft.LAST_NAME = LFt.LAST_NAME
      Ft.LONG_NAME = LFt.LONG_NAME
      Ft.MASTER_FLAG = LFt.MASTER_FLAG
      Ft.MIDDLE_NAME = LFt.MIDDLE_NAME
      Ft.SHORT_NAME = LFt.SHORT_NAME
      LFt.PREFIX_ID = LFt.PREFIX_ID
      LFt.EMAIL = LFt.EMAIL
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateName = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateName = False
End Function

Private Function MigrateEnterprise() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CEnterprise
Dim LFt As CLegacyEnterprise
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "ENTERPRISE"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyEnterprise
   LFt.ENTERPRISE_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CEnterprise
      Ft.AddEditMode = SHOW_ADD
      Ft.ENTERPRISE_ID = LFt.ENTERPRISE_ID
      Ft.BUSINESS_TYPE = LFt.BUSINESS_TYPE
      Ft.BRANCH_CODE = ""
      Ft.EMAIL = LFt.EMAIL
      Ft.ENTERPRISE_TYPE = -1
      Ft.POLICY = LFt.POLICY
      Ft.TAX_ID = LFt.TAX_ID
      Ft.WEBSITE = LFt.WEBSITE
      Ft.SETUP_DATE = LFt.SETUP_DATE
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateEnterprise = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateEnterprise = False
End Function

Private Function MigrateAddress() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CAddress
Dim LFt As CLegacyAddress
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "ADDRESS"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyAddress
   LFt.ADDRESS_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CAddress
      Ft.AddEditMode = SHOW_ADD
      Ft.ADDRESS_ID = LFt.ADDRESS_ID
      Ft.ADDRESS_TYPE = LFt.ADDRESS_TYPE
      Ft.AMPHUR = LFt.AMPHUR
      Ft.BANGKOK_FLAG = LFt.BANGKOK_FLAG
      Ft.COUNTRY_ID = LFt.COUNTRY_ID
      Ft.FAX1 = LFt.FAX1
      Ft.FAX2 = LFt.FAX2
      Ft.HOME = LFt.HOME
      Ft.MOO = LFt.MOO
      Ft.PHONE1 = LFt.PHONE1
      Ft.PHONE2 = LFt.PHONE2
      Ft.PROVINCE = LFt.PROVINCE
      Ft.ROAD = LFt.ROAD
      Ft.SOI = LFt.SOI
      Ft.VILLAGE = LFt.VILLAGE
      Ft.ZIPCODE = LFt.ZIPCODE
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateAddress = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateAddress = False
End Function

Private Function MigrateEnpName() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CEnterpriseName
Dim LFt As CLegacyEnpName
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "ENTERPRISE_NAME"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyEnpName
   LFt.ENTERPRISE_NAME_ID = -1
   LFt.ENTERPRISE_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CEnterpriseName
      Ft.AddEditMode = SHOW_ADD
      Ft.ENTERPRISE_NAME_ID = LFt.ENTERPRISE_NAME_ID
      Ft.ENTERPRISE_ID = LFt.ENTERPRISE_ID
      Ft.NAME_ID = LFt.NAME_ID
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateEnpName = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateEnpName = False
End Function

Private Function MigrateEnpAddress() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CEnterpriseAddress
Dim LFt As CLegacyEnpAddr
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "ENTERPRISE_ADDRESS"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyEnpAddr
   LFt.ENTERPRISE_ADDRESS_ID = -1
   LFt.ENTERPRISE_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CEnterpriseAddress
      Ft.AddEditMode = SHOW_ADD
      Ft.ENTERPRISE_ADDRESS_ID = LFt.ENTERPRISE_ADDRESS_ID
      Ft.ENTERPRISE_ID = LFt.ENTERPRISE_ID
      Ft.ADDRESS_ID = LFt.ADDRESS_ID
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateEnpAddress = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateEnpAddress = False
End Function

Private Function MigrateEnpPerson() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CEnterprisePerson
Dim LFt As CLegacyEnpPerson
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "ENTERPRISE_PERSON"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyEnpPerson
   LFt.ENTERPRISE_PERSON_ID = -1
   LFt.ENTERPRISE_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CEnterprisePerson
      Ft.AddEditMode = SHOW_ADD
      Ft.ENTERPRISE_PERSON_ID = LFt.ENTERPRISE_PERSON_ID
      Ft.ENTERPRISE_ID = LFt.ENTERPRISE_ID
      Ft.NAME_ID = LFt.NAME_ID
      Ft.MASTER_FLAG = LFt.MASTER_FLAG
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateEnpPerson = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateEnpPerson = False
End Function

Private Function MigrateAccount() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CAccount
Dim LFt As CLegacyAccount
Dim iCount As Long
Dim I As Long
Dim MaxId As Long

   MaxId = 0
   
   txtTableName.Text = "ACCOUNT"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyAccount
   LFt.ACCOUNT_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CAccount
      Ft.AddEditMode = SHOW_ADD
      Ft.ACCOUNT_ID = LFt.ACCOUNT_ID
      Ft.ACCOUNT_NO = LFt.ACCOUNT_NO
      Ft.ACCOUNT_STATUS = LFt.ACCOUNT_STATUS
      Ft.ACCOUNT_TYPE = LFt.ACCOUNT_TYPE
      Ft.CUSTOMER_ID = LFt.CUSTOMER_ID
      Ft.NOTE = LFt.NOTE
      Ft.Credit = LFt.Credit
      Ft.ENABLE_FLAG = LFt.ENABLE_FLAG
      Ft.MASTER_FLAG = LFt.MASTER_FLAG
      Ft.CUSTOMER_ID = LFt.CUSTOMER_ID
      
      If Ft.ACCOUNT_ID > MaxId Then
         MaxId = Ft.ACCOUNT_ID
      End If
      
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateAccount = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateAccount = False
End Function

Private Function MigrateSubscriber() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CSubscriber
Dim LFt As CLegacySubscriber
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "SUBSCRIBER"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacySubscriber
   LFt.SUBSCRIBER_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CSubscriber
      Ft.AddEditMode = SHOW_ADD
      Ft.SUBSCRIBER_ID = LFt.SUBSCRIBER_ID
      Ft.ACCOUNT_ID = LFt.ACCOUNT_ID
      Ft.DUMMY_FLAG = LFt.DUMMY_FLAG
      Ft.SUBSCRIBER_DESC = LFt.SUBSCRIBER_DESC
      Ft.SUBSCRIBER_NO = LFt.SUBSCRIBER_NO
      Ft.SUBSCRIBER_STATUS = LFt.SUBSCRIBER_STATUS
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateSubscriber = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateSubscriber = False
End Function

Private Function MigrateCustomer() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CCustomer
Dim LFt As CLegacyCustomer
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "CUSTOMER"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyCustomer
   LFt.CUSTOMER_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CCustomer
      Ft.AddEditMode = SHOW_ADD
      Ft.CUSTOMER_ID = LFt.CUSTOMER_ID
      Ft.CUSTOMER_CODE = LFt.CUSTOMER_CODE
      Ft.CUSTOMER_GRADE = LFt.CUSTOMER_GRADE
      Ft.Credit = LFt.Credit
      Ft.NORMAL_DISCOUNT = LFt.NORMAL_DISCOUNT
      Ft.TAX_ID = LFt.TAX_ID
      Ft.CUSTOMER_TYPE = -1
      Ft.EMAIL = LFt.EMAIL
      Ft.WEBSITE = LFt.WEBSITE
      Ft.BIRTH_DATE = LFt.BIRTH_DATE
      Ft.CUSTOMER_PASSWORD = LFt.CUSTOMER_PASSWORD
      Ft.BUSINESS_TYPE = LFt.BUSINESS_TYPE
      Ft.BUSINESS_DESC = LFt.BUSINESS_DESC
      Ft.RESPONSE_BY = LFt.RESPONSE_BY
      
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateCustomer = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateCustomer = False
End Function

Private Function MigrateCustomerName() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CCustomerName
Dim LFt As CLegacyCustomerName
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "CUSTOMER_NAME"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyCustomerName
   LFt.CUSTOMER_NAME_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CCustomerName
      Ft.AddEditMode = SHOW_ADD
      Ft.CUSTOMER_NAME_ID = LFt.CUSTOMER_NAME_ID
      Ft.CUSTOMER_ID = LFt.CUSTOMER_ID
      Ft.NAME_ID = LFt.NAME_ID
      
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateCustomerName = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateCustomerName = False
End Function

Private Function MigrateCustomerAddress() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CCustomerAddress
Dim LFt As CLegacyCustAddr
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "CUSTOMER_ADDRESS"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyCustAddr
   LFt.CUSTOMER_ADDRESS_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CCustomerAddress
      Ft.AddEditMode = SHOW_ADD
      Ft.CUSTOMER_ADDRESS_ID = LFt.CUSTOMER_ADDRESS_ID
      Ft.CUSTOMER_ID = LFt.CUSTOMER_ID
      Ft.ADDRESS_ID = LFt.ADDRESS_ID
      Ft.ADDRESS_TYPE = LFt.ADDRESS_TYPE

      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateCustomerAddress = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateCustomerAddress = False
End Function

Private Function MigrateAgreement() As Boolean
On Error GoTo ErrorHandler
Dim Ft As CAgreement
Dim LFt As CLegacyAgreement
Dim iCount As Long
Dim I As Long

   txtTableName.Text = "AGREEMENT"
   prgProgress.MIN = 0
   prgProgress.Value = 0
   prgProgress.MAX = 100
   Me.Refresh
   
   Set LFt = New CLegacyAgreement
   LFt.AGREEMENT_ID = -1
   Call LFt.QueryData(m_Rs, iCount)
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      Call LFt.PopulateFromRS(m_Rs)
      
      Set Ft = New CAgreement
      Ft.AddEditMode = SHOW_ADD
      Ft.AGREEMENT_ID = LFt.AGREEMENT_ID
      Ft.AGREEMENT_ID = LFt.AGREEMENT_ID
      Ft.SOC_FEATURE_ID = LFt.SOC_FEATURE_ID
      Ft.SUBSCRIBER_ID = LFt.SUBSCRIBER_ID
      Ft.EXCLUDE_FLAG = LFt.EXCLUDE_FLAG
      Ft.EFFECTIVE_DATE = LFt.EFFECTIVE_DATE
      Ft.EXPIRE_DATE = LFt.EXPIRE_DATE
      Ft.ISSUE_DATE = LFt.ISSUE_DATE
      Ft.SOC_ID = LFt.SOC_ID
      
      Call Ft.AddEditData(False)
      Set Ft = Nothing
      
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      m_Rs.MoveNext
   Wend
   
   Set LFt = Nothing
   MigrateAgreement = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   
   MigrateAgreement = False
End Function

Private Sub cmdStart_Click()
On Error GoTo ErrorHandler
Dim HasBegin As Boolean
Dim MaxSeq As Long

   Call EnableForm(Me, False)
   
   HasBegin = False
   Call glbDatabaseMngr.ConnectLegacyDatabase(txtFileName.Text, glbParameterObj.UserName, glbParameterObj.Password, glbErrorLog)
   Call glbDaily.StartTransaction
   HasBegin = True
   
   If Not MigrateCountry Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("COUNTRY_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("COUNTRY_SEQ", MaxSeq)
   
   If Not MigrateFeatureType Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("FEATURE_TYPE_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("FEATURE_TYPE_SEQ", MaxSeq)
   
   If Not MigrateUnit Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("UNIT_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("UNIT_SEQ", MaxSeq)
   
   If Not MigrateFeature Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("FEATURE_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("FEATURE_SEQ", MaxSeq)
   
   If Not MigrateSoc Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("SOC_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("SOC_SEQ", MaxSeq)
   
   If Not MigrateSocFeature Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("SOC_FEATURE_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("SOC_FEATURE_SEQ", MaxSeq)
   
   If Not MigrateStpTierVol Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("STPTIER_VOL_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("STPTIER_VOL_SEQ", MaxSeq)
   
   If Not MigrateAcRate Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("AC_RATE_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("AC_RATE_SEQ", MaxSeq)
   
   If Not MigrateOcRate Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("OC_RATE_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("OC_RATE_SEQ", MaxSeq)
   
   If Not MigrateRcRate Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("RC_RATE_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("RC_RATE_SEQ", MaxSeq)
   
   If Not MigrateUcRate Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("UC_RATE_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("UC_RATE_SEQ", MaxSeq)
   
   If Not MigrateName Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("NAME_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("NAME_SEQ", MaxSeq)
   
   If Not MigrateAddress Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("ADDRESS_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("ADDRESS_SEQ", MaxSeq)
   
   If Not MigrateEnterprise Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("ENTERPRISE_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("ENTERPRISE_SEQ", MaxSeq)
   
   If Not MigrateEnpName Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("ENTERPRISE_NAME_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("ENTERPRISE_NAME_SEQ", MaxSeq)
   
   If Not MigrateEnpAddress Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("ENTERPRISE_ADDRESS_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("ENTERPRISE_ADDRESS_SEQ", MaxSeq)
   
   If Not MigrateEnpPerson Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("ENTERPRISE_PERSON_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("ENTERPRISE_PERSON_SEQ", MaxSeq)
   
   If Not MigrateCustomer Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("CUSTOMER_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("CUSTOMER_SEQ", MaxSeq)
   
   If Not MigrateCustomerName Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("CUSTOMER_NAME_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("CUSTOMER_NAME_SEQ", MaxSeq)
   
   If Not MigrateCustomerAddress Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("CUSTOMER_ADDRESS_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("CUSTOMER_ADDRESS_SEQ", MaxSeq)
   
   If Not MigrateAccount() Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("ACCOUNT_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("ACCOUNT_SEQ", MaxSeq)
   
   If Not MigrateSubscriber Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call glbDatabaseMngr.GetLegacySeqID("SUBSCRIBER_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("SUBSCRIBER_SEQ", MaxSeq)
   
   If Not MigrateAgreement Then
      Call glbDaily.RollbackTransaction
      Call glbDatabaseMngr.DisConnectLegacyDatabase
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call glbDatabaseMngr.GetLegacySeqID("AGREEMENT_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("AGREEMENT_SEQ", MaxSeq)
   
   Call glbDatabaseMngr.GetLegacySeqID("RECEIPT_NUMBER_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("RECEIPT_NUMBER_SEQ", MaxSeq)
   
   Call glbDatabaseMngr.GetLegacySeqID("DO_NUMBER_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("DO_NUMBER_SEQ", MaxSeq)
   
   Call glbDatabaseMngr.GetLegacySeqID("PO_NUMBER_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("PO_NUMBER_SEQ", MaxSeq)
   
   Call glbDatabaseMngr.GetLegacySeqID("CUSTOMER_CODE_SEQ", MaxSeq, glbErrorLog, 0)
   Call glbDatabaseMngr.SetSeqID("CUSTOMER_NUMBER_SEQ", MaxSeq)
   
   Call glbDaily.CommitTransaction
   HasBegin = False
   Call glbDatabaseMngr.DisConnectLegacyDatabase
   Call EnableForm(Me, True)
   
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      Call glbDaily.RollbackTransaction
   End If
   Call glbDatabaseMngr.DisConnectLegacyDatabase
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         id = 0
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

Private Sub ResetStatus()
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "ปรับราคาเฉลี่ย"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "ชื่อไฟล์")
   Call InitNormalLabel(lblTableName, "ชื่อตาราง")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTableName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtTableName.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFile.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFile, MapText("..."))
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Employee = New CEmployee
   Set m_Rs = New ADODB.Recordset
   
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

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub


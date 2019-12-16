VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditSupplier 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditSupplier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
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
         TabIndex        =   7
         Top             =   2880
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
         TabIndex        =   6
         Top             =   2880
         Width           =   3495
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   9
         Top             =   4170
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
         TabIndex        =   3
         Top             =   1530
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtShortName 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtEmail 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1980
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWebSite 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   2430
         Width           =   3465
         _ExtentX        =   16960
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBusinessDesc 
         Height          =   450
         Left            =   1860
         TabIndex        =   8
         Top             =   3330
         Width           =   9225
         _ExtentX        =   16907
         _ExtentY        =   794
      End
      Begin prjFarmManagement.uctlTextBox txtCredit 
         Height          =   435
         Left            =   5700
         TabIndex        =   1
         Top             =   1080
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin prjFarmManagement.uctlTextBox txtDiscountPercent 
         Height          =   435
         Left            =   8070
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3015
         Left            =   150
         TabIndex        =   10
         Top             =   4710
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5318
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
         Column(1)       =   "frmAddEditSupplier.frx":27A2
         Column(2)       =   "frmAddEditSupplier.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditSupplier.frx":290E
         FormatStyle(2)  =   "frmAddEditSupplier.frx":2A6A
         FormatStyle(3)  =   "frmAddEditSupplier.frx":2B1A
         FormatStyle(4)  =   "frmAddEditSupplier.frx":2BCE
         FormatStyle(5)  =   "frmAddEditSupplier.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditSupplier.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtSupplierChequeName 
         Height          =   435
         Left            =   6990
         TabIndex        =   28
         Top             =   2400
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   767
      End
      Begin VB.Label lblSupplierChequeName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5400
         TabIndex        =   29
         Top             =   2490
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSupplier.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   15
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSupplier.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSupplier.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblDiscountPercent 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6900
         TabIndex        =   26
         Top             =   1140
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   6570
         TabIndex        =   25
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4710
         TabIndex        =   24
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label lblBusinessDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   23
         Top             =   3450
         Width           =   1695
      End
      Begin VB.Label lblWebsite 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   22
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   21
         Top             =   2070
         Width           =   1575
      End
      Begin VB.Label lblShortName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         TabIndex        =   20
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label lblEnterpriseType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5400
         TabIndex        =   19
         Top             =   2940
         Width           =   1485
      End
      Begin VB.Label lblBusinessType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   18
         Top             =   2940
         Width           =   1485
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   1620
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddEditSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Supplier As CSupplier
Private m_PartItems As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long

Private FileName As String

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_Supplier.SUPPLIER_ID = id
      If Not glbDaily.QuerySupplier(m_Supplier, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Supplier.PopulateFromRS(1, m_Rs)
      
      txtEmail.Text = m_Supplier.EMAIL
      txtWebSite.Text = m_Supplier.WEBSITE
      cboBusinessType.ListIndex = IDToListIndex(cboBusinessType, m_Supplier.SUPPLIER_TYPE)
      cboEnterpriseType.ListIndex = IDToListIndex(cboEnterpriseType, m_Supplier.SUPPLIER_GRADE)
      txtShortName.Text = m_Supplier.SUPPLIER_CODE
      txtBusinessDesc.Text = m_Supplier.BUSINESS_DESC
      txtCredit.Text = m_Supplier.Credit
      txtSupplierChequeName.Text = m_Supplier.SUPPLIER_CHEQUE_NAME
      
            
      Dim NAME As CName
      Dim CstName As CSupplierName
      If (Not m_Supplier.CstNames Is Nothing) And (m_Supplier.CstNames.Count > 0) Then
         Set CstName = m_Supplier.CstNames(1)
         Set NAME = CstName.NAME
         txtName.Text = NAME.LONG_NAME
      Else
         txtName.Text = ""
      End If
   Else
      ShowMode = SHOW_ADD
   End If
   
   If ShowMode = SHOW_ADD Then
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

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("MAIN_SUPPLIER_EDIT") Then
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

   If Not CheckUniqueNs(SUPPLIER_UNIQUE, txtShortName.Text, id) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtShortName.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   m_Supplier.AddEditMode = ShowMode
   m_Supplier.BIRTH_DATE = -1
   m_Supplier.EMAIL = txtEmail.Text
   m_Supplier.WEBSITE = txtWebSite.Text
   m_Supplier.SUPPLIER_TYPE = cboBusinessType.ItemData(Minus2Zero(cboBusinessType.ListIndex))
   m_Supplier.SUPPLIER_GRADE = cboEnterpriseType.ItemData(Minus2Zero(cboEnterpriseType.ListIndex))
   m_Supplier.Credit = Val(txtCredit.Text)
   m_Supplier.SUPPLIER_CODE = txtShortName.Text
   m_Supplier.BUSINESS_DESC = txtBusinessDesc.Text
   m_Supplier.SUPPLIER_CHEQUE_NAME = txtSupplierChequeName.Text
   
   Dim CstName As CSupplierName
   If m_Supplier.CstNames.Count <= 0 Then
      Set CstName = New CSupplierName
      CstName.Flag = "A"
      Call m_Supplier.CstNames.add(CstName)
   Else
      Set CstName = m_Supplier.CstNames.Item(1)
      CstName.Flag = "E"
   End If
   
   Dim NAME As CName
   If m_Supplier.CstNames.Count <= 0 Then
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
   If Not glbDaily.AddEditSupplier(m_Supplier, IsOK, True, glbErrorLog) Then
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

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditSupplierAddress.TempCollection = m_Supplier.CstAddr
      frmAddEditSupplierAddress.ShowMode = SHOW_ADD
      frmAddEditSupplierAddress.HeaderText = MapText("เพิ่มที่อยู่")
      Load frmAddEditSupplierAddress
      frmAddEditSupplierAddress.Show 1

      OKClick = frmAddEditSupplierAddress.OKClick

      Unload frmAddEditSupplierAddress
      Set frmAddEditSupplierAddress = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Supplier.CstAddr)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      Set frmAddEditContactPerson.TempCollection = m_Supplier.CstContacts
      frmAddEditContactPerson.ShowMode = SHOW_ADD
      frmAddEditContactPerson.HeaderText = MapText("เพิ่มข้อมูลผู้ติดต่อ")
      Load frmAddEditContactPerson
      frmAddEditContactPerson.Show 1

      OKClick = frmAddEditContactPerson.OKClick

      Unload frmAddEditContactPerson
      Set frmAddEditContactPerson = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Supplier.CstContacts)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      Set frmAddEditSupplierSpec.TempCollection = m_Supplier.SupplierSpecs
      frmAddEditSupplierSpec.ShowMode = SHOW_ADD
      frmAddEditSupplierSpec.HeaderText = MapText("เพิ่มข้อมูลผู้เสปค")
      Load frmAddEditSupplierSpec
      frmAddEditSupplierSpec.Show 1

      OKClick = frmAddEditSupplierSpec.OKClick

      Unload frmAddEditSupplierSpec
      Set frmAddEditSupplierSpec = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Supplier.SupplierSpecs)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      Set frmAddEditSupplierUsed.TempCollection = m_Supplier.SupplierUseds
      frmAddEditSupplierUsed.ShowMode = SHOW_ADD
      frmAddEditSupplierUsed.HeaderText = MapText("เพิ่มข้อมูลวัตถุดิบ")
      Load frmAddEditSupplierUsed
      frmAddEditSupplierUsed.Show 1

      OKClick = frmAddEditSupplierUsed.OKClick

      Unload frmAddEditSupplierUsed
      Set frmAddEditSupplierUsed = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Supplier.SupplierUseds)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      Set frmAddEditSupplierTransport.TempCollection = m_Supplier.SupplierTranSport
      frmAddEditSupplierTransport.ShowMode = SHOW_ADD
      frmAddEditSupplierTransport.HeaderText = MapText("เพิ่มข้อมูลรถขนส่ง")
      Load frmAddEditSupplierTransport
      frmAddEditSupplierTransport.Show 1

      OKClick = frmAddEditSupplierTransport.OKClick

      Unload frmAddEditSupplierTransport
      Set frmAddEditSupplierTransport = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Supplier.SupplierTranSport)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 7 Then
      Set frmAddEditSupplierAccount.TempCollection = m_Supplier.SupplierAccount
      frmAddEditSupplierAccount.ShowMode = SHOW_ADD
      frmAddEditSupplierAccount.HeaderText = MapText("เพิ่มข้อมูลบัญชีธนาคาร")
      Load frmAddEditSupplierAccount
      frmAddEditSupplierAccount.Show 1

      OKClick = frmAddEditSupplierAccount.OKClick

      Unload frmAddEditSupplierAccount
      Set frmAddEditSupplierAccount = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Supplier.SupplierAccount)
         GridEX1.Rebind
      End If
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(4)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_Supplier.CstAddr.Remove (ID2)
      Else
         m_Supplier.CstAddr.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Supplier.CstAddr)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      If ID1 <= 0 Then
         m_Supplier.CstContacts.Remove (ID2)
      Else
         m_Supplier.CstContacts.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Supplier.CstContacts)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If ID1 <= 0 Then
         m_Supplier.SupplierSpecs.Remove (ID2)
      Else
         m_Supplier.SupplierSpecs.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Supplier.SupplierSpecs)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      If ID1 <= 0 Then
         m_Supplier.SupplierUseds.Remove (ID2)
      Else
         m_Supplier.SupplierUseds.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Supplier.SupplierUseds)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      If ID1 <= 0 Then
         m_Supplier.SupplierTranSport.Remove (ID2)
      Else
         m_Supplier.SupplierTranSport.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Supplier.SupplierTranSport)
      GridEX1.Rebind
      m_HasModify = True
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   id = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditSupplierAddress.id = id
      Set frmAddEditSupplierAddress.TempCollection = m_Supplier.CstAddr
      frmAddEditSupplierAddress.HeaderText = MapText("แก้ไขที่อยู่")
      frmAddEditSupplierAddress.ShowMode = SHOW_EDIT
      Load frmAddEditSupplierAddress
      frmAddEditSupplierAddress.Show 1

      OKClick = frmAddEditSupplierAddress.OKClick

      Unload frmAddEditSupplierAddress
      Set frmAddEditSupplierAddress = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Supplier.CstAddr)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      frmAddEditContactPerson.id = id
      Set frmAddEditContactPerson.TempCollection = m_Supplier.CstContacts
      frmAddEditContactPerson.HeaderText = MapText("แก้ไขข้อมูลผู้ติดต่อ")
      frmAddEditContactPerson.ShowMode = SHOW_EDIT
      Load frmAddEditContactPerson
      frmAddEditContactPerson.Show 1

      OKClick = frmAddEditContactPerson.OKClick

      Unload frmAddEditContactPerson
      Set frmAddEditContactPerson = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Supplier.CstContacts)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      frmAddEditSupplierSpec.id = id
      Set frmAddEditSupplierSpec.TempCollection = m_Supplier.SupplierSpecs
      frmAddEditSupplierSpec.HeaderText = MapText("แก้ไขข้อมูลเสปค")
      frmAddEditSupplierSpec.ShowMode = SHOW_EDIT
      Load frmAddEditSupplierSpec
      frmAddEditSupplierSpec.Show 1

      OKClick = frmAddEditSupplierSpec.OKClick

      Unload frmAddEditSupplierSpec
      Set frmAddEditSupplierSpec = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Supplier.SupplierSpecs)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      frmAddEditSupplierUsed.id = id
      Set frmAddEditSupplierUsed.TempCollection = m_Supplier.SupplierUseds
      frmAddEditSupplierUsed.HeaderText = MapText("แก้ไขข้อมูลวัตถุดิบ")
      frmAddEditSupplierUsed.ShowMode = SHOW_EDIT
      Load frmAddEditSupplierUsed
      frmAddEditSupplierUsed.Show 1

      OKClick = frmAddEditSupplierUsed.OKClick

      Unload frmAddEditSupplierUsed
      Set frmAddEditSupplierUsed = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Supplier.SupplierUseds)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      frmAddEditSupplierTransport.id = id
      Set frmAddEditSupplierTransport.TempCollection = m_Supplier.SupplierTranSport
      frmAddEditSupplierTransport.HeaderText = MapText("แก้ไขข้อมูลรถขนส่ง")
      frmAddEditSupplierTransport.ShowMode = SHOW_EDIT
      Load frmAddEditSupplierTransport
      frmAddEditSupplierTransport.Show 1

      OKClick = frmAddEditSupplierTransport.OKClick

      Unload frmAddEditSupplierTransport
      Set frmAddEditSupplierTransport = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Supplier.SupplierTranSport)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 7 Then
      frmAddEditSupplierAccount.id = id
      Set frmAddEditSupplierAccount.TempCollection = m_Supplier.SupplierAccount
      frmAddEditSupplierAccount.HeaderText = MapText("แก้ไขข้อมูลเลขที่บัญชี")
      frmAddEditSupplierAccount.ShowMode = SHOW_EDIT
      Load frmAddEditSupplierAccount
      frmAddEditSupplierAccount.Show 1

      OKClick = frmAddEditSupplierAccount.OKClick

      Unload frmAddEditSupplierAccount
      Set frmAddEditSupplierAccount = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Supplier.SupplierAccount)
         GridEX1.Rebind
      End If
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
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
      Call LoadSupplierType(cboBusinessType)
      Call LoadSupplierGrade(cboEnterpriseType)
'      Call LoadPartItem(Nothing, m_PartItems, , "")
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Supplier.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_Supplier.QueryFlag = 0
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
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   TabStrip1.Width = GridEX1.Width
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
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
   
   Set m_Supplier = Nothing
   Set m_PartItems = Nothing
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
   Col.Width = 2370
   Col.Caption = MapText("รหัสวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 6360
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2550
   Col.Caption = MapText("บาร์โค้ด")
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
   Col.Width = 2370
   Col.Caption = MapText("รหัสวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 6360
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2550
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("% ความชื้น")
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
   Col.Visible = False
   Col.Caption = MapText("ID")
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("Real ID")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2835
   Col.Caption = MapText("ชื่อ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2745
   Col.Caption = MapText("นามสกุล")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2535
   Col.Caption = MapText("อีเมลล์")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 3450
   Col.Caption = MapText("ตำแหน่ง")
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
   Col.Width = 2370
   Col.Caption = MapText("รหัสวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 6360
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2550
   Col.Caption = MapText("สถานะ")
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
   Col.Width = 2370
   Col.Caption = MapText("รหัสทะเบียนรถ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 6360
   Col.Caption = MapText("เลขทะเบียนรถ")

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
   Col.Width = 1500
   Col.Caption = MapText("หมายเลขบัญชี")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("ชื่อบัญชี")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("ชื่อธนาคาร")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("สาขา")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1200
   Col.Caption = MapText("แสดงในค่าขนส่ง")

End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblWebsite, MapText("เว็บไซต์"))
   Call InitNormalLabel(lblShortName, MapText("รหัสซัพ ฯ"))
   Call InitNormalLabel(lblEnterpriseType, MapText("ระดับซัพ ฯ"))
   Call InitNormalLabel(lblName, MapText("ชื่อซัพ ฯ"))
   Call InitNormalLabel(lblEmail, MapText("อีเมลล์"))
   Call InitNormalLabel(lblBusinessType, MapText("ประเภทซัพ ฯ"))
   Call InitNormalLabel(lblBusinessDesc, MapText("รายละเอียดซัพ ฯ"))
   Call InitNormalLabel(lblCredit, MapText("เครดิต"))
   Call InitNormalLabel(Label2, MapText("วัน"))
   Call InitNormalLabel(lblDiscountPercent, MapText("% ส่วนลด"))
   Call InitNormalLabel(lblSupplierChequeName, MapText("ชื่อพิมพ์เช็ค"))
   
   Call InitCombo(cboBusinessType)
   Call InitCombo(cboEnterpriseType)
   
   Call txtShortName.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtEmail.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtWebSite.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBusinessDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtCredit.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
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
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("ที่อยู่")
   TabStrip1.Tabs.add().Caption = MapText("วัตถุดิบรับเข้า")
   TabStrip1.Tabs.add().Caption = MapText("ผู้ติดต่อ")
   TabStrip1.Tabs.add().Caption = MapText("ข้อมูลจำเพาะ")
   TabStrip1.Tabs.add().Caption = MapText("วัตถุดิบ")
   TabStrip1.Tabs.add().Caption = MapText("รถขนส่ง")
   TabStrip1.Tabs.add().Caption = MapText("บัญชีธนาคาร")
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
   Set m_Supplier = New CSupplier
   Set m_PartItems = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
'   If TabStrip1.SelectedItem.Index = 1 Then
'      RowBuffer.RowStyle = RowBuffer.Value(7)
'   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_Supplier.CstAddr Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CSupplierAddress
      Dim Addr As CAddress
      If m_Supplier.CstAddr.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Supplier.CstAddr, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      Set Addr = CR.Addresses

      Values(1) = Addr.ADDRESS_ID
      Values(2) = RealIndex
      Values(3) = Addr.PackAddress
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_Supplier.PartItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Li As CLotItem
      Dim Pi As CPartItem
      If m_Supplier.PartItems.Count <= 0 Then
         Exit Sub
      End If
      Set Li = GetItem(m_Supplier.PartItems, RowIndex, RealIndex)
      If Li Is Nothing Then
         Exit Sub
      End If

      Values(1) = Li.PART_ITEM_ID
      Values(2) = RealIndex
      Values(3) = Li.PART_NO
      Values(4) = Li.PART_DESC
      Values(5) = ""
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      If m_Supplier.CstContacts Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CC As CSupplierContact
      Dim N As CName
      If m_Supplier.CstContacts.Count <= 0 Then
         Exit Sub
      End If
      Set CC = GetItem(m_Supplier.CstContacts, RowIndex, RealIndex)
      If CC Is Nothing Then
         Exit Sub
      End If
      Set N = CC.NAME

      Values(1) = N.NAME_ID
      Values(2) = RealIndex
      Values(3) = N.LONG_NAME
      Values(4) = N.LAST_NAME
      Values(5) = N.EMAIL
      Values(6) = CC.CONTACT_POSITION
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If m_Supplier.SupplierSpecs Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Sp As CSupplierSpec
      If m_Supplier.SupplierSpecs.Count <= 0 Then
         Exit Sub
      End If
      Set Sp = GetItem(m_Supplier.SupplierSpecs, RowIndex, RealIndex)
      If Sp Is Nothing Then
         Exit Sub
      End If

      Values(1) = Sp.SUPPLIER_SPEC_ID
      Values(2) = RealIndex
      Values(3) = Sp.PART_NO
      Values(4) = Sp.PART_DESC
      Values(5) = FormatNumber(Sp.HUMIDITY)
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      If m_Supplier.SupplierUseds Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Su As CSupplierUsed
      If m_Supplier.SupplierUseds.Count <= 0 Then
         Exit Sub
      End If
      Set Su = GetItem(m_Supplier.SupplierUseds, RowIndex, RealIndex)
      If Su Is Nothing Then
         Exit Sub
      End If

      Values(1) = Su.SUPPLIER_USED_ID
      Values(2) = RealIndex
      Values(3) = Su.PART_NO
      Values(4) = Su.PART_DESC
      Values(5) = Su.USED_FLAG
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      If m_Supplier.SupplierTranSport Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim St As CSupplierTranSport
      If m_Supplier.SupplierTranSport.Count <= 0 Then
         Exit Sub
      End If
      Set St = GetItem(m_Supplier.SupplierTranSport, RowIndex, RealIndex)
      If St Is Nothing Then
         Exit Sub
      End If

      Values(1) = St.SUPPLIER_TRANSPORT_ID
      Values(2) = RealIndex
      Values(3) = St.SUPPLIER_TRANSPORT_CODE
      Values(4) = St.SUPPLIER_TRANSPORT_DETAIL
   ElseIf TabStrip1.SelectedItem.Index = 7 Then
      If m_Supplier.SupplierAccount Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim SA As CSupplierAccount
      If m_Supplier.SupplierAccount.Count <= 0 Then
         Exit Sub
      End If
      Set SA = GetItem(m_Supplier.SupplierAccount, RowIndex, RealIndex)
      If SA Is Nothing Then
         Exit Sub
      End If

      Values(1) = SA.SUPPLIER_ACCOUNT_ID
      Values(2) = RealIndex
      Values(3) = SA.SUPPLIER_ACCOUNT_NO
      Values(4) = SA.SUPPLIER_ACCOUNT_NAME
      Values(5) = SA.SUPPLIER_ACCOUNT_BANK
      Values(6) = SA.SUPPLIER_ACCOUNT_BRANCH
      Values(7) = IIf(SA.USE_TRANSPORT_FLAG = "N", "", "ใช้งาน")
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_Supplier.CstAddr)
      GridEX1.Rebind
   
      cmdAdd.Enabled = True
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid2
      GridEX1.ItemCount = CountItem(m_Supplier.PartItems)
      GridEX1.Rebind
      
      cmdAdd.Enabled = False
      cmdEdit.Enabled = False
      cmdDelete.Enabled = False
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      Call InitGrid3
      GridEX1.ItemCount = CountItem(m_Supplier.CstContacts)
      GridEX1.Rebind
   
      cmdAdd.Enabled = True
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      Call InitGrid4
      GridEX1.ItemCount = CountItem(m_Supplier.SupplierSpecs)
      GridEX1.Rebind
   
      cmdAdd.Enabled = True
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      Call InitGrid5
      GridEX1.ItemCount = CountItem(m_Supplier.SupplierUseds)
      GridEX1.Rebind
   
      cmdAdd.Enabled = True
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      Call InitGrid6
      GridEX1.ItemCount = CountItem(m_Supplier.SupplierTranSport)
      GridEX1.Rebind
   
      cmdAdd.Enabled = True
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
   ElseIf TabStrip1.SelectedItem.Index = 7 Then
      Call InitGrid7
      GridEX1.ItemCount = CountItem(m_Supplier.SupplierAccount)
      GridEX1.Rebind
   
      cmdAdd.Enabled = True
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
   End If
End Sub

Private Sub txtBusinessDesc_Change()
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

Private Sub txtDiscountPercent_Change()
   m_HasModify = True
End Sub

Private Sub txtEmail_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtShortName_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtShortName_LostFocus()
   If Not CheckUniqueNs(SUPPLIER_UNIQUE, txtShortName.Text, id) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtShortName.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
End Sub

Private Sub txtSupplierChequeName_Change()
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

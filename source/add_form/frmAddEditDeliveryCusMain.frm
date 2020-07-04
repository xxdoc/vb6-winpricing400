VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditDeliveryCusMain 
   BackColor       =   &H80000000&
   ClientHeight    =   9450
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11970
   Icon            =   "frmAddEditDeliveryCusMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   11970
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   9495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   16748
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   0
         Top             =   1560
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
      Begin GridEX20.GridEX GridEX1 
         Height          =   6615
         Left            =   120
         TabIndex        =   1
         Top             =   2040
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   11668
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
         Column(1)       =   "frmAddEditDeliveryCusMain.frx":27A2
         Column(2)       =   "frmAddEditDeliveryCusMain.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditDeliveryCusMain.frx":290E
         FormatStyle(2)  =   "frmAddEditDeliveryCusMain.frx":2A6A
         FormatStyle(3)  =   "frmAddEditDeliveryCusMain.frx":2B1A
         FormatStyle(4)  =   "frmAddEditDeliveryCusMain.frx":2BCE
         FormatStyle(5)  =   "frmAddEditDeliveryCusMain.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditDeliveryCusMain.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   2100
         TabIndex        =   10
         Top             =   960
         Width           =   7785
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblCustomerLookup 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   9
         Top             =   1050
         Width           =   1995
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3390
         TabIndex        =   4
         Top             =   8790
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDeliveryCusMain.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   120
         TabIndex        =   2
         Top             =   8790
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDeliveryCusMain.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1740
         TabIndex        =   3
         Top             =   8790
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   6
         Top             =   8790
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   5
         Top             =   8790
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDeliveryCusMain.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditDeliveryCusMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Customers As Collection
Private m_DeliveryCus As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public CustomerID As Long


Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim DC As CDeliveryCus
'
'   If Not VerifyTextControl(lblName, txtName, False) Then
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
'   m_Enterprise.AddEditMode = ShowMode
'   m_Enterprise.SETUP_DATE = -1
'   m_Enterprise.TAX_ID = txtTaxID.Text
'   m_Enterprise.EMAIL = ""
'   m_Enterprise.POLICY = ""
'   m_Enterprise.WEBSITE = ""
'   m_Enterprise.BUSINESS_TYPE = -1
'   m_Enterprise.ENTERPRISE_TYPE = -1
'   m_Enterprise.BRANCH_CODE = txtBranchCode.Text
'   m_Enterprise.BRANCH_NAME = txtBranchName.Text
'
'   Dim EnpName As CEnterpriseName
'   If m_Enterprise.EnpNames.Count <= 0 Then
'      Set EnpName = New CEnterpriseName
'      EnpName.Flag = "A"
'      Call m_Enterprise.EnpNames.add(EnpName)
'   Else
'      Set EnpName = m_Enterprise.EnpNames.Item(1)
'      EnpName.Flag = "E"
'   End If
'
'   Dim NAME As CName
'   If EnpName.Names.Count <= 0 Then
'      Set EnpName.Names = New Collection
'      Set NAME = New CName
'      NAME.LONG_NAME = txtName.Text
'      NAME.SHORT_NAME = txtShortName.Text
'      NAME.Flag = "A"
'      Call EnpName.Names.add(NAME)
'      Set NAME = Nothing
'   Else
'      Set NAME = EnpName.Names.Item(1)
'      NAME.LONG_NAME = txtName.Text
'      NAME.SHORT_NAME = txtShortName.Text
'      NAME.Flag = "E"
'   End If

   Call EnableForm(Me, False)
   glbDatabaseMngr.DBConnection.BeginTrans
           For Each DC In m_DeliveryCus
               DC.DELIVERY_CUS_ITEM_CODE = DC.DELIVERY_CUS_ITEM_CODE
               DC.DELIVERY_CUS_ITEM_NAME = DC.DELIVERY_CUS_ITEM_NAME
               DC.CUSTOMER_ID = CustomerID
                  
               If DC.Flag = "D" Then
                   Call DC.DeleteData
               ElseIf DC.Flag = "A" Then
                  DC.AddEditMode = SHOW_ADD
               ElseIf DC.Flag = "E" Then
                 DC.AddEditMode = SHOW_EDIT
               End If
               
               If DC.Flag = "A" Or DC.Flag = "E" Then
                  If Not glbDaily.AddEditDeliveryCus(DC, IsOK, False, glbErrorLog) Then
                      Call EnableForm(Me, True)
                   End If
                    If Not IsOK Then
                         Call EnableForm(Me, True)
                         glbErrorLog.ShowUserError
                         glbDatabaseMngr.DBConnection.RollbackTrans
                     End If
                  End If
             Next DC
   Call EnableForm(Me, True)
   glbDatabaseMngr.DBConnection.CommitTrans
   SaveData = True
End Function

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
       Set frmAddEditDeliveryCus.ParentForm = Me
      Set frmAddEditDeliveryCus.TempCollection = m_DeliveryCus
      frmAddEditDeliveryCus.ShowMode = SHOW_ADD
      frmAddEditDeliveryCus.HeaderText = MapText("เพิ่มข้อมูลสถานที่จัดส่ง")
      frmAddEditDeliveryCus.CustomerCode = uctlCustomerLookup.MyTextBox.Text
      Load frmAddEditDeliveryCus
      frmAddEditDeliveryCus.Show 1

      OKClick = frmAddEditDeliveryCus.OKClick

      Unload frmAddEditDeliveryCus
      Set frmAddEditDeliveryCus = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_DeliveryCus)
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
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
         If ID1 <= 0 Then
         m_DeliveryCus.Remove (ID2)
      Else
         m_DeliveryCus.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_DeliveryCus)
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
      frmAddEditDeliveryCus.id = id
      frmAddEditDeliveryCus.ID2 = Val(GridEX1.Value(1))
      Set frmAddEditDeliveryCus.ParentForm = Me
      Set frmAddEditDeliveryCus.TempCollection = m_DeliveryCus
      frmAddEditDeliveryCus.ShowMode = SHOW_EDIT
      frmAddEditDeliveryCus.HeaderText = MapText("แก้ไขสถานที่จัดส่ง")
      Load frmAddEditDeliveryCus
      frmAddEditDeliveryCus.Show 1

      OKClick = frmAddEditDeliveryCus.OKClick

      Unload frmAddEditDeliveryCus
      Set frmAddEditDeliveryCus = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_DeliveryCus)
         GridEX1.Rebind
      End If
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
    
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      QueryData (True)
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If

End Sub

Private Sub cmdSearch_Click()

End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)

      Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
      Set uctlCustomerLookup.MyCollection = m_Customers
      
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, CustomerID)
      
      Call QueryData(True)
      m_HasModify = False
      Call EnableForm(Me, True)
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
   
   Set m_Customers = Nothing
   Set m_DeliveryCus = Nothing
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
   Col.Width = 2355
   Col.Caption = MapText("รหัสสถานที่")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 5000
   Col.Caption = MapText("ชื่อสถานที่จัดส่ง")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("สถานะ")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลสถานที่จัดส่ง")
   pnlHeader.Caption = MapText("ข้อมูลสถานที่จัดส่ง")
   
   Call InitNormalLabel(lblCustomerLookup, MapText("ชื่อลูกค้า"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
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
   TabStrip1.Tabs.add().Caption = MapText("ข้อมูลสถานที่จัดส่ง")
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
   
   Set m_Customers = New Collection
   Set m_DeliveryCus = New Collection
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
     If m_DeliveryCus Is Nothing Then
         Exit Sub
      End If
      
      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim DC As CDeliveryCus
      If m_DeliveryCus.Count <= 0 Then
         Exit Sub
      End If
      Set DC = GetItem(m_DeliveryCus, RowIndex, RealIndex)
      If DC Is Nothing Then
         Exit Sub
      End If

      Values(1) = DC.DELIVERY_CUS_ITEM_ID
      Values(2) = RealIndex
      Values(3) = DC.DELIVERY_CUS_ITEM_CODE
      Values(4) = DC.DELIVERY_CUS_ITEM_NAME
      Values(5) = IIf(DC.HIDE_FLAG = "N", "ใช้งาน", "ยกเลิก")
      
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_DeliveryCus)
      GridEX1.Rebind
   End If
End Sub

Private Sub txtEmail_Change()
   m_HasModify = True
End Sub

Private Sub txtBranchCode_Change()
   m_HasModify = True
End Sub

Private Sub txtBranchName_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtShortName_Change()
   m_HasModify = True
End Sub

Private Sub txtSlogan_Change()
   m_HasModify = True
End Sub

Private Sub txtTaxID_Change()
   m_HasModify = True
End Sub


Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True

   If Flag Then
      Call EnableForm(Me, False)
      Call LoadDeliveryCus(Nothing, m_DeliveryCus, CustomerID, , , , "") 'LOAD สถานที่จัดส่ง
   End If
   
   GridEX1.ItemCount = CountItem(m_DeliveryCus)
   GridEX1.Rebind
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub
Private Sub uctlCustomerLookup_Change()
Dim Customer As CCustomer

   CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
'   If CustomerID > 0 Then
      Call LoadDeliveryCus(Nothing, m_DeliveryCus, CustomerID, , , , "N") 'LOAD สถานที่จัดส่ง
'   End If
   
   Call TabStrip1_Click
   m_HasModify = True
End Sub

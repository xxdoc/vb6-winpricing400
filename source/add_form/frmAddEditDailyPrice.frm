VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditDailyPrice 
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8925
   Icon            =   "frmAddEditDailyPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   8925
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4380
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   7726
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlSellTypeLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1410
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlPriceDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   990
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtSalePrice 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1860
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBuyPrice 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   2310
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlEmployeeLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   2760
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblEmployee 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   15
         Top             =   2790
         Width           =   1455
      End
      Begin VB.Label lblSellType 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   1470
         Width           =   1605
      End
      Begin VB.Label lblBuyPrice 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   13
         Top             =   2430
         Width           =   1695
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   3300
         TabIndex        =   12
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3300
         TabIndex        =   11
         Top             =   2400
         Width           =   765
      End
      Begin VB.Label lblPriceDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   1020
         Width           =   1575
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2955
         TabIndex        =   5
         Top             =   3510
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDailyPrice.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4605
         TabIndex        =   6
         Top             =   3510
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblSalePrice 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   8
         Top             =   1980
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAddEditDailyPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_InventoryDoc As CInventoryDoc
Private m_Employees As Collection
Private m_SellTypes As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private FileName As String
Private m_SumUnit As Double

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
'      m_InventoryDoc.INVENTORY_DOC_ID = ID
'      m_InventoryDoc.COMMIT_FLAG = ""
'      If Not glbDaily.QueryInventoryDoc(m_InventoryDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
'         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
   End If
ItemCount = 0 'Remove this when done
   If ItemCount > 0 Then
'      Call m_InventoryDoc.PopulateFromRS(1, m_Rs)
'
'      uctlPriceDate.ShowDate = m_InventoryDoc.DOCUMENT_DATE
'      txtSalePrice.Text = Format(m_InventoryDoc.DELIVERY_FEE, "0.00")
'      uctlSellTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlSellTypeLookup.MyCombo, m_InventoryDoc.SUPPLIER_ID)
'      uctlEmployeeLookup.MyCombo.ListIndex = IDToListIndex(uctlEmployeeLookup.MyCombo, m_InventoryDoc.DELIVERY_ID)
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
   
   If Not VerifyDate(lblPriceDate, uctlPriceDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblSellType, uctlSellTypeLookup.MyCombo, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblSalePrice, txtSalePrice, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblBuyPrice, txtBuyPrice, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
'
'   m_InventoryDoc.AddEditMode = ShowMode
'   m_InventoryDoc.INVENTORY_DOC_ID = ID
'    m_InventoryDoc.DOCUMENT_DATE = uctlPriceDate.ShowDate
'   m_InventoryDoc.DO_NO = txtDoNo.Text
'   m_InventoryDoc.TRUCK_NO = txtTruckNo.Text
'   m_InventoryDoc.DOCUMENT_NO = txtDocumentNo.Text
'   m_InventoryDoc.DELIVERY_FEE = Val(txtSalePrice.Text)
'   m_InventoryDoc.BILL_NO = txtDeliveryNo.Text
'   m_InventoryDoc.SENDER_NAME = txtSender.Text
'   m_InventoryDoc.RECEIVE_NAME = txtReceiver.Text
'   m_InventoryDoc.SUPPLIER_ID = uctlSellTypeLookup.MyCombo.ItemData(Minus2Zero(uctlSellTypeLookup.MyCombo.ListIndex))
'   m_InventoryDoc.DELIVERY_ID = uctlEmployeeLookup.MyCombo.ItemData(Minus2Zero(uctlEmployeeLookup.MyCombo.ListIndex))
'   m_InventoryDoc.DOCUMENT_TYPE = 1
'   m_InventoryDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
      
   Call EnableForm(Me, False)
'   If Not glbDaily.AddEditInventoryDoc(m_InventoryDoc, IsOK, True, glbErrorLog) Then
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'      SaveData = False
'      Call EnableForm(Me, True)
'      Exit Function
'   End If
'   If Not IsOK Then
'      Call EnableForm(Me, True)
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
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

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadEmployee(uctlEmployeeLookup.MyCombo, m_Employees)
      Set uctlEmployeeLookup.MyCollection = m_Employees
      
      Call LoadSellType(uctlSellTypeLookup.MyCombo, m_SellTypes)
      Set uctlSellTypeLookup.MyCollection = m_SellTypes
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_InventoryDoc.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_InventoryDoc.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_InventoryDoc = Nothing
   Set m_Employees = Nothing
   Set m_SellTypes = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption

   Call InitNormalLabel(lblSalePrice, MapText("�ҤҢ��"))
   Call InitNormalLabel(lblPriceDate, MapText("�Ҥ� � �ѹ���"))
   Call InitNormalLabel(lblBuyPrice, MapText("�Ҥҫ���"))
   Call InitNormalLabel(Label1, MapText("�ҷ"))
   Call InitNormalLabel(Label2, MapText("�ҷ"))
   Call InitNormalLabel(lblSellType, MapText("�������Ҥҷͧ"))
   Call InitNormalLabel(lblEmployee, MapText("������¡��"))
   
   Call txtSalePrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtBuyPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)

   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
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
   Set m_InventoryDoc = New CInventoryDoc
   Set m_Employees = New Collection
   Set m_SellTypes = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeliveryNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtTruckNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPriceDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlSellTypeLookup_Change()
   m_HasModify = True
End Sub
